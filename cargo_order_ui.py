"""Small button editor; reuses the dispatch fragment and does not fetch data."""
from hashlib import sha1
from cargo_order import label, validate


def editor(st, key, initial, capacity, rerun, *, disabled=False):
    st.session_state.setdefault(key, list(initial))
    order = list(st.session_state[key])
    # Scope mobile overrides to this editor, never to other app columns.
    container_key = 'cargo_compact_' + sha1(key.encode()).hexdigest()[:12]
    selector = '.st-key-' + container_key
    st.markdown(f'''<style>
    {selector} [data-testid="stHorizontalBlock"] {{
        flex-direction:row !important; flex-wrap:nowrap !important;
        gap:6px !important; align-items:center !important;
    }}
    {selector} [data-testid="stColumn"] {{
        min-width:0 !important; width:auto !important;
        flex:1 1 0% !important;
    }}
    {selector} [data-testid="stVerticalBlock"] {{gap:5px !important;}}
    {selector} [data-testid="stColumn"]:first-child {{flex:3 1 0% !important;}}
    {selector} .st-key-{container_key}_actions [data-testid="stColumn"] {{flex:1 1 0% !important;}}
    {selector} button {{
        width:100% !important; min-height:44px !important; height:44px !important;
        padding:4px 8px !important; margin:0 !important; border-radius:10px !important;
        box-shadow:none !important; font-size:14px !important;
    }}
    {selector} button p {{font-size:14px !important; white-space:nowrap; margin:0 !important;}}
    {selector} [data-testid="stCaptionContainer"] p {{font-size:12px !important; margin:0 !important;}}
    </style>''', unsafe_allow_html=True)
    with st.container(key=container_key):
        st.caption(f'車頭 ↓ 車尾先卸 · 一般 {order.count("bike")}／電 {order.count("ebike")} · {len(order)}/{capacity}')
        for i, kind in enumerate(order):
            cols = st.columns([3, 1, 1])
            if cols[0].button(f'{i + 1:02d}　' + ('一般車' if kind == 'bike' else '電車') + (' · 車頭' if i == 0 else '') + (' · 車尾' if i == len(order) - 1 else ''), key=f'{key}:type:{i}', disabled=disabled):
                order[i] = 'ebike' if kind == 'bike' else 'bike'
                st.session_state[key] = order
                rerun()
            for col, delta, title in ((cols[1], -1, '↑'), (cols[2], 1, '↓')):
                j = i + delta
                if col.button(title, key=f'{key}:move:{i}:{delta}', disabled=disabled or not 0 <= j < len(order)):
                    order[i], order[j] = order[j], order[i]
                    st.session_state[key] = order
                    rerun()
        with st.container(key=container_key + "_actions"):
            cols = st.columns(3)
            for col, kind, title in ((cols[0], 'bike', '＋一般'), (cols[1], 'ebike', '＋電車')):
                if col.button(title, key=f'{key}:add:{kind}', disabled=disabled or len(order) >= capacity):
                    st.session_state[key] = order + [kind]
                    rerun()
            if cols[2].button('移除車尾', key=f'{key}:pop', disabled=disabled or not order):
                st.session_state[key] = order[:-1]
                rerun()
    return order



def recover_locked_cargo(st, prefix, active_trip, capacity):
    """Restore a confirmed, locked trip snapshot after an iPhone/WebSocket reconnect.

    A draft is never authoritative: the trip's cargo_before is the sequence
    actually accepted when this destination was locked. Reject corrupt snapshots.
    """
    if not isinstance(active_trip, dict):
        return False
    state = st.session_state
    enabled_key = prefix + '::cargo_enabled'
    if 'cargo_before' not in active_trip:
        # A trip accepted in ordinary mode must not inherit a stale cargo toggle.
        if state.get(enabled_key):
            state[enabled_key] = False
            state[prefix + '::cargo_was_enabled'] = False
        return False
    try:
        snapshot = validate(active_trip['cargo_before'], capacity)
    except ValueError:
        return False
    if (snapshot.count('bike') != active_trip.get('truck_before_bike')
            or snapshot.count('ebike') != active_trip.get('truck_before_ebike')):
        return False
    restored = {
        enabled_key: True,
        prefix + '::cargo_was_enabled': True,
        prefix + '::cargo_order': snapshot,
        prefix + '::cargo_ready': True,
        prefix + '::truck_bike': snapshot.count('bike'),
        prefix + '::truck_ebike': snapshot.count('ebike'),
    }
    for key, value in restored.items():
        if state.get(key) != value:
            state[key] = value
    # An unconfirmed draft may outlive the confirmation widget, but cannot
    # invalidate the already-locked trip or block completing/cancelling it.
    state.pop(prefix + '::cargo_draft', None)
    return True


def settings(st, prefix, capacity, locked, rerun):
    order_key = prefix + '::cargo_order'
    ready_key = prefix + '::cargo_ready'
    previous_key = prefix + '::cargo_was_enabled'
    enabled = st.checkbox('依車上排列調度', key=prefix + '::cargo_enabled', disabled=locked,
                          help='單列、只從車尾卸車；關閉時使用原本智慧調度。')
    if enabled != st.session_state.get(previous_key, False):
        st.session_state[ready_key] = False
        st.session_state.pop(prefix + '::cargo_draft', None)
    st.session_state[previous_key] = enabled
    if not enabled:
        return None, True
    ready = st.session_state.get(ready_key, False)
    raw_current = st.session_state.get(order_key, [])
    try:
        current = validate(raw_current, capacity)
    except ValueError:
        current = []
        ready = False
        st.session_state[ready_key] = False
    # A genuinely empty truck has no loading sequence to confirm. Do not
    # auto-confirm a nonempty draft or a stale nonempty saved order.
    if (not locked and not ready and isinstance(raw_current, (list, tuple))
            and not raw_current and not st.session_state.get(prefix + '::cargo_draft')
            and st.session_state.get(prefix + '::truck_bike', 0) == 0
            and st.session_state.get(prefix + '::truck_ebike', 0) == 0):
        st.session_state[order_key] = []
        st.session_state[ready_key] = True
        ready = True
    with st.expander('車上排列設定' + ('（請確認）' if not ready else ''), expanded=not ready):
        if locked:
            st.write(label(current))
            st.caption('本站已鎖定；本次上車順序可在下方調整。')
        else:
            draft = editor(st, prefix + '::cargo_draft', current, capacity, rerun)
            if st.button('確認目前排列', key=prefix + '::cargo_confirm', type='primary'):
                try:
                    current = validate(draft, capacity)
                except ValueError as exc:
                    st.error(str(exc))
                else:
                    st.session_state[order_key] = current
                    st.session_state[ready_key] = True
                    st.session_state[prefix + '::truck_bike'] = current.count('bike')
                    st.session_state[prefix + '::truck_ebike'] = current.count('ebike')
                    rerun()
    return list(current), bool(ready and (locked or st.session_state.get(prefix + "::cargo_draft", current) == current))


def rerun_fragment(st):
    # User interaction normally runs inside the existing dispatch fragment.
    # Fall back on first full render, where fragment-only rerun is not allowed.
    from streamlit.errors import StreamlitAPIException
    try:
        st.rerun(scope="fragment")
    except StreamlitAPIException:
        st.rerun()
