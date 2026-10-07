import ast
import importlib.util
from pathlib import Path
import sys
import threading
import types
import unittest
from unittest.mock import Mock, patch

from background_refresh import BackgroundRefresh, RefreshState, live_refresh

ROOT = Path(__file__).resolve().parents[1]


class LiveRecoveryTests(unittest.TestCase):
    def test_cold_catalog_and_parking_result_is_not_discarded(self):
        now = [0.0]
        release = threading.Event()
        manager = BackgroundRefresh(timeout=live_refresh.timeout, clock=lambda: now[0])
        def loader():
            release.wait(2)
            return {"records": ["fresh"]}
        try:
            manager.poll("scope", loader)
            now[0] = 66  # catalog retries + parking round, previously discarded at 25s
            self.assertTrue(manager.poll("scope", loader).pending)
            release.set()
            manager._jobs["scope"].future.result(timeout=2)
            state = manager.poll("scope", loader)
            self.assertEqual(state.value, {"records": ["fresh"]})
            self.assertEqual(state.error, "")
        finally:
            release.set()
            manager._pool.shutdown(wait=True)

    def test_all_legacy_compatibility_patches_still_apply(self):
        # Production app.py is native-integrated. Keep validating the archived
        # patcher only as a migration regression fixture.
        tree = ast.parse((ROOT / "app_patch_legacy.py").read_text())
        ns = {"source": (ROOT / "legacy_ui.py").read_text(), "BATTERY_ICON_DATA_URI": "fixture"}
        active = False
        for node in tree.body:
            if isinstance(node, ast.FunctionDef) and node.name == "replace_exact":
                active = True
            if not active:
                continue
            take = isinstance(node, ast.FunctionDef) and node.name in ("replace_exact", "_modernize_legacy_iframes")
            take |= isinstance(node, ast.Assign)
            take |= (isinstance(node, ast.Expr) and isinstance(node.value, ast.Call)
                     and isinstance(node.value.func, ast.Name) and node.value.func.id == "replace_exact")
            if take:
                exec(compile(ast.Module(body=[node], type_ignores=[]), "app.py", "exec"), ns)
        compile(ns["source"], "patched_legacy.py", "exec")
        self.assertIn("browser_fallback=browser_fallback", ns["source"])

    def load_ui(self):
        st = types.ModuleType("streamlit")
        st.session_state = {}
        st.fragment = lambda **kwargs: lambda fn: fn
        for name in ("caption", "warning", "info", "iframe", "rerun"):
            setattr(st, name, Mock())
        spec = importlib.util.spec_from_file_location("recovery_ui", ROOT / "live_status_ui.py")
        ui = importlib.util.module_from_spec(spec)
        with patch.dict(sys.modules, {"streamlit": st}):
            spec.loader.exec_module(ui)
        return ui, st

    def test_error_mounts_browser_once_and_recovery_returns_server(self):
        ui, st = self.load_ui()
        ui.live_refresh = Mock()
        failed = RefreshState(None, False, "HTTP 503", 1, None)
        ui.live_refresh.poll.return_value = failed
        fallback = Mock(return_value={"ok": True, "event_id": "browser", "records": [1]})
        stations = [{"name": "測試站"}]
        self.assertEqual(ui.render_background_live_status(stations, browser_fallback=fallback)["event_id"], "browser")
        st.rerun.assert_not_called()
        key = next(k for k in st.session_state if not k.endswith("::error"))
        ui._monitor(ui.station_scope(stations), key)
        st.rerun.assert_not_called()
        ready = RefreshState({"ok": True, "event_id": "server", "records": [2]}, False, "", 2, 0)
        ui.live_refresh.poll.return_value = ready
        ui._monitor(ui.station_scope(stations), key)
        st.rerun.assert_called_once()
        result = ui.render_background_live_status(stations, browser_fallback=fallback)
        self.assertEqual(result["event_id"], "server")
        fallback.assert_called_once()

    def test_first_async_error_requests_full_rerun_for_fallback(self):
        ui, st = self.load_ui()
        ui.live_refresh = Mock()
        ui.live_refresh.poll.return_value = RefreshState(None, False, "offline", 1, None)
        ui._monitor((("測試站", ""),), "scope")
        ui._monitor((("測試站", ""),), "scope")
        st.rerun.assert_called_once()


if __name__ == "__main__":
    unittest.main()
