from __future__ import annotations

"""Compatibility shim for Streamlit v1 components declared from exec() code.

The legacy UI is executed through ``exec(compile(...))`` by the V29 compatibility
entrypoint. Streamlit's v1 ``declare_component`` inspects its direct caller to
build a component name. Code executed from the legacy compile context may not
map to an importable Python module, which makes ``inspect.getmodule`` return
None and causes Streamlit to fail before the component is registered.

Installing this shim routes declarations through a real imported module while
preserving Streamlit's original API and behavior.
"""

from typing import Any

import streamlit.components.v1 as components


_ORIGINAL_DECLARE_COMPONENT = components.declare_component
_INSTALLED = False


def _module_safe_declare_component(
    name: str,
    path: str | None = None,
    url: str | None = None,
) -> Any:
    """Call Streamlit's original declaration from a normal imported module."""
    return _ORIGINAL_DECLARE_COMPONENT(name=name, path=path, url=url)


def install_component_declare_compat() -> None:
    """Install once for all legacy v1 custom-component declarations."""
    global _INSTALLED
    if _INSTALLED:
        return
    components.declare_component = _module_safe_declare_component
    _INSTALLED = True
