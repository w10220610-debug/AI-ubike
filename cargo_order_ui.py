"""Small button editor; reuses the dispatch fragment and does not fetch data."""
from cargo_order import label, validate


def editor(st, key, initial, capacity, rerun, *, disabled=False):
    st.session_state.setdefault(key, list(initial))
    order = list(st.session_state[key])
    st.caption('由車頭排到車尾；新車加在車尾。點車種可切換，↑↓可交換位置。')
    st.write(label(order))
    for i, kind in enumerate(order):
        cols = st.columns([3, 1, 1])
        if cols[0].button(f'{i + 1}｜' + ('一般車' if kind == 'bike' else '電車'), key=f'{key}:type:{i}', disabled=disabled):
            order[i] = 'ebike' if kind == 'bike' else 'bike'
            st.session_state[key] = order
            rerun()
        for col, delta, title in ((cols[1], -1, '↑'), (cols[2], 1, '↓')):
            j = i + delta
            if col.button(title, key=f'{key}:move:{i}:{delta}', disabled=disabled or not 0 <= j < len(order)):
                order[i], order[j] = order[j], order[i]
                st.session_state[key] = order
                rerun()
    cols = st.columns(3)
    for col, kind, title in ((cols[0], 'bike', '＋一般車'), (cols[1], 'ebike', '＋電車')):
        if col.button(title, key=f'{key}:add:{kind}', disabled=disabled or len(order) >= capacity):
            st.session_state[key] = order + [kind]
            rerun()
    if cols[2].button('移除車尾', key=f'{key}:pop', disabled=disabled or not order):
        st.session_state[key] = order[:-1]
        rerun()
    return order


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
    current = st.session_state.get(order_key, [])
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
    return list(current), ready and st.session_state.get(prefix + "::cargo_draft", current) == current


def rerun_fragment(st):
    # User interaction normally runs inside the existing dispatch fragment.
    # Fall back on first full render, where fragment-only rerun is not allowed.
    from streamlit.errors import StreamlitAPIException
    try:
        st.rerun(scope="fragment")
    except StreamlitAPIException:
        st.rerun()
