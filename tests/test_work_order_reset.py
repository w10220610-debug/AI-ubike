from pathlib import Path
from unittest.mock import patch
import time
import pandas as pd
from streamlit.testing.v1 import AppTest
from work_orders import clear_draft, lookup_draft, merge, orders, pending


def test_clear_draft_keeps_saved_orders_and_other_inputs():
    state = {'p::text': 'wrong', 'p::draft': [{}], 'p::editor::1': {},
             'p::draft_version': '1', 'p::notice': 'old', 'p::event': 'old-event',
             'other::text': 'keep', 'saved_orders': [{}]}
    clear_draft(state, 'p')
    assert state == {'p::event': 'old-event', 'other::text': 'keep',
                     'saved_orders': [{}], 'p::reset_epoch': 1}
    clear_draft(state, 'p')
    assert state['p::reset_epoch'] == 2


def test_removed_order_does_not_return_from_stale_browser_backup():
    cache = {}
    row = dict(station_name='A', kind='換電', equipment='0123456', updated_at=1)
    merge(cache, [row])
    orders(cache)[0].update(state='deleted', updated_at=2)
    assert merge(cache, [row]) == 0
    assert not pending(cache)
    restored = {}
    merge(restored, orders(cache))
    assert not pending(restored)


def test_vehicle_review_retains_leading_zeros_and_rejects_unknown_station():
    row = dict(station_name='A', bike_no='0123456', pillar_no='03', queried_at='2026/10/6 22:10')
    draft = lookup_draft([row, row, dict(row, station_name='UNKNOWN')], ['A'], '換電')
    assert len(draft) == 1
    assert draft[0]['equipment'] == '0123456'
    assert '03' in draft[0]['description']
    assert lookup_draft([row], ['A'], 'invalid') == []


def test_real_streamlit_reset_after_failed_parse_and_remove():
    def fixture():
        import streamlit as st
        import pandas as pd
        from work_orders import render_manager, merge
        cache = st.session_state.setdefault('cache', {})
        if 'seeded' not in st.session_state:
            merge(cache, [dict(station_name='A', kind='換電', equipment='0123456')])
            st.session_state['seeded'] = True
        render_manager(pd.DataFrame({'場站名稱': ['A']}), cache=cache, token='fixture',
                       page_mode='一般分析', save=lambda: None)
    with patch('work_orders.lookup_station_args', return_value=dict(stations=[], catalog_pending=False, catalog_error='')):
        app = AppTest.from_function(fixture).run()
        app.text_area[0].set_value('沒有場站的錯誤輸入').run()
        next(x for x in app.button if x.label == '整理為待確認清單').click().run()
        assert app.session_state['work_orders::fixture::draft'] == []
        next(x for x in app.button if '清除輸入' in x.label).click().run()
        assert not app.exception
        assert app.text_area[0].value == ''
        assert len(pending(app.session_state['cache'])) == 1
        next(x for x in app.button if x.label == '移除誤加派工').click().run()
        assert not app.exception
        assert not pending(app.session_state['cache'])
