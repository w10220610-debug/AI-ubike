from __future__ import annotations

"""One-time builder that turns the V29 compatibility patcher into a native app.py.

The source of truth for the migration is app_patch_legacy.py (or app.py on the
first run) plus legacy_ui.py. Every replace_exact() patch is applied at build
-time, then the support functions are placed directly ahead of the integrated
legacy UI. The generated app.py contains no runtime source rewriting or exec().
"""

import ast
from pathlib import Path
import shutil
import sys


ROOT = Path(__file__).resolve().parents[1]
PATCHER_PATH = ROOT / "app_patch_legacy.py"
CURRENT_APP_PATH = ROOT / "app.py"
LEGACY_PATH = ROOT / "legacy_ui.py"
OUTPUT_PATH = ROOT / "app.py"


def _node_text(source: str, node: ast.AST) -> str:
    lines = source.splitlines(keepends=True)
    start = int(getattr(node, "lineno", 1)) - 1
    end = int(getattr(node, "end_lineno", start + 1))
    return "".join(lines[start:end])


def _assigned_names(node: ast.AST) -> set[str]:
    names: set[str] = set()
    if isinstance(node, ast.Assign):
        targets = node.targets
    elif isinstance(node, ast.AnnAssign):
        targets = [node.target]
    else:
        return names
    for target in targets:
        if isinstance(target, ast.Name):
            names.add(target.id)
        elif isinstance(target, (ast.Tuple, ast.List)):
            for item in target.elts:
                if isinstance(item, ast.Name):
                    names.add(item.id)
    return names


def _is_install_compat_call(node: ast.AST) -> bool:
    if not isinstance(node, ast.Expr) or not isinstance(node.value, ast.Call):
        return False
    func = node.value.func
    return isinstance(func, ast.Name) and func.id == "install_component_declare_compat"


def _is_replace_call(node: ast.AST) -> bool:
    if not isinstance(node, ast.Expr) or not isinstance(node.value, ast.Call):
        return False
    func = node.value.func
    return isinstance(func, ast.Name) and func.id == "replace_exact"


def _extract_support_source(patcher_source: str, tree: ast.Module) -> str:
    replace_def_line = None
    for node in tree.body:
        if isinstance(node, ast.FunctionDef) and node.name == "replace_exact":
            replace_def_line = node.lineno
            break
    if replace_def_line is None:
        raise RuntimeError("找不到 replace_exact()；來源似乎已不是 V29 patcher。")

    chunks: list[str] = []
    for node in tree.body:
        if getattr(node, "lineno", 10**9) >= replace_def_line:
            break
        if _assigned_names(node) & {"LEGACY_APP", "source"}:
            continue
        if _is_install_compat_call(node):
            continue
        chunks.append(_node_text(patcher_source, node).rstrip() + "\n\n")

    chunks.append(
        "# Streamlit 1.59 legacy custom-component compatibility.\n"
        "install_component_declare_compat()\n\n"
        'FORMAL_APP_ARCHITECTURE = "native-integrated-v29"\n\n'
    )
    return "".join(chunks)


def _static_env(tree: ast.Module) -> dict[str, object]:
    sys.path.insert(0, str(ROOT))
    from battery_icon_data import BATTERY_ICON_DATA_URI

    env: dict[str, object] = {"BATTERY_ICON_DATA_URI": BATTERY_ICON_DATA_URI}
    for node in tree.body:
        if not isinstance(node, ast.Assign) or len(node.targets) != 1:
            continue
        target = node.targets[0]
        if not isinstance(target, ast.Name):
            continue
        try:
            env[target.id] = _eval_static(node.value, env)
        except Exception:
            continue
    return env


def _eval_static(node: ast.AST, env: dict[str, object]):
    if isinstance(node, ast.Constant):
        return node.value
    if isinstance(node, ast.Name) and node.id in env:
        return env[node.id]
    if isinstance(node, ast.BinOp) and isinstance(node.op, ast.Add):
        return _eval_static(node.left, env) + _eval_static(node.right, env)
    if isinstance(node, ast.JoinedStr):
        parts: list[str] = []
        for value in node.values:
            if isinstance(value, ast.Constant):
                parts.append(str(value.value))
            elif isinstance(value, ast.FormattedValue):
                parts.append(str(_eval_static(value.value, env)))
            else:
                raise ValueError("unsupported f-string node")
        return "".join(parts)
    if isinstance(node, ast.Call) and isinstance(node.func, ast.Attribute):
        base = _eval_static(node.func.value, env)
        args = [_eval_static(arg, env) for arg in node.args]
        if node.func.attr == "replace" and isinstance(base, str):
            return base.replace(*args)
    raise ValueError(f"unsupported static expression: {ast.dump(node, include_attributes=False)}")


def _apply_replace_patches(patcher_source: str, tree: ast.Module, legacy_source: str) -> tuple[str, int]:
    patch_count = 0
    env = _static_env(tree)
    for node in tree.body:
        if not _is_replace_call(node):
            continue
        call = node.value
        if len(call.args) < 2:
            raise RuntimeError(f"replace_exact at line {node.lineno} has too few args")
        try:
            old = _eval_static(call.args[0], env)
            new = _eval_static(call.args[1], env)
        except Exception as exc:
            raise RuntimeError(
                f"replace_exact at line {node.lineno} is not statically evaluable: {exc}"
            ) from exc
        if not isinstance(old, str) or not isinstance(new, str):
            raise RuntimeError(f"replace_exact at line {node.lineno} must use strings")
        label = f"line {node.lineno}"
        for keyword in call.keywords:
            if keyword.arg == "label":
                try:
                    label = str(ast.literal_eval(keyword.value))
                except Exception:
                    pass
        count = legacy_source.count(old)
        if count != 1:
            raise RuntimeError(
                f"{label}: expected exactly one match, found {count}"
            )
        legacy_source = legacy_source.replace(old, new, 1)
        patch_count += 1
    return legacy_source, patch_count


def _modernize_iframes(source: str) -> str:
    source = source.replace("components.html(", "st.iframe(")
    source = source.replace("height=0", "height=1")
    source = source.replace("scrolling=False,", "")
    source = source.replace(", scrolling=False", "")
    return source


def _strip_duplicate_future_import(source: str) -> str:
    lines = source.splitlines(keepends=True)
    return "".join(
        line for line in lines
        if line.strip() != "from __future__ import annotations"
    ).lstrip("\n")


def build() -> None:
    source_path = PATCHER_PATH if PATCHER_PATH.exists() else CURRENT_APP_PATH
    patcher_source = source_path.read_text(encoding="utf-8")
    legacy_source = LEGACY_PATH.read_text(encoding="utf-8")

    if not PATCHER_PATH.exists():
        shutil.copy2(CURRENT_APP_PATH, PATCHER_PATH)

    tree = ast.parse(patcher_source)
    support = _extract_support_source(patcher_source, tree)
    integrated, patch_count = _apply_replace_patches(
        patcher_source, tree, legacy_source
    )
    integrated = _modernize_iframes(integrated)
    integrated = _strip_duplicate_future_import(integrated)

    header = (
        "# AUTO-GENERATED ONCE BY tools/build_formal_app.py\n"
        "# V29 formal architecture: native integrated runtime.\n"
        "# No runtime legacy_ui source patching, no compile/exec compatibility layer.\n\n"
    )
    output = header + support + integrated

    forbidden = (
        "LEGACY_APP.read_text",
        "exec(compile_legacy_source",
        "def replace_exact(",
    )
    for token in forbidden:
        if token in output:
            raise RuntimeError(f"formal app still contains forbidden runtime patch token: {token}")

    compile(output, str(OUTPUT_PATH), "exec")
    OUTPUT_PATH.write_text(output, encoding="utf-8")

    print(
        f"Built formal app.py: {patch_count} patches integrated, "
        f"{len(output.splitlines())} lines"
    )


if __name__ == "__main__":
    build()

# Workflow trigger marker: formal migration.
