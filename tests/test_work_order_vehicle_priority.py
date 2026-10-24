"""Regression cases transcribed from the supplied dispatch screenshots."""
import unittest
from work_orders import parse_text, apply_lookup_to_draft, prepare_vehicle_draft, DRAFT_SCHEMA


class VehiclePriorityTests(unittest.TestCase):
    def test_station_code_is_not_vehicle_and_old_station_is_not_routed(self):
        rows = parse_text('電池 2.0E 501501080 1560268 車輛電池 電量低於70%\n臺東縣長期照顧中心\n電池 2.0E 1590102 車輛電池', ['臺東縣長期照顧中心', '新站'])
        self.assertEqual([r['equipment'] for r in rows], ['1560268', '1590102'])
        self.assertTrue(all(not r['station_name'] and not r['selected'] for r in rows))
        matched = apply_lookup_to_draft(rows, [{'bike_no': '1560268', 'station_name': '新站', 'pillar_no': '12'}], ['新站'], '換電')
        self.assertEqual(matched[0]['station_name'], '新站')
        self.assertIn('臺東縣長期照顧中心', matched[0]['description'])
        self.assertEqual(matched[0]['kind'], '電池')
        self.assertFalse(matched[1]['selected'])
        self.assertEqual(matched[1]['station_name'], '')

    def test_all_sixteen_vehicle_numbers(self):
        ids = '1560268 1590102 1590129 1590176 1560301 1590021 1590031 1590039 1590047 1590244 1590094 1590103 1560173 1560192 1590232 1560269'.split()
        text = '\n'.join('電池 2.0E 501501091 ' + bike + ' 車輛電池-正常營運中車輛' for bike in ids)
        rows = parse_text(text, [])
        self.assertEqual([r['equipment'] for r in rows], ids)
        self.assertTrue(all(not r['station_name'] for r in rows))

    def test_preserve_leading_zero_and_existing_station_work(self):
        rows = parse_text('舊站 車輛故障 0123456\n新站 清潔', ['舊站', '新站'])
        self.assertEqual(rows[0]['equipment'], '0123456')
        self.assertEqual(rows[0]['source_station'], '舊站')
        self.assertEqual(rows[0]['station_name'], '')
        self.assertEqual(rows[1]['station_name'], '新站')
        self.assertTrue(rows[1]['selected'])

    def test_multiple_numbers_and_bad_lookup(self):
        rows = parse_text('1560268,1590102', [])
        self.assertEqual(len(rows), 2)
        out = apply_lookup_to_draft(rows, [{'bike_no': '1560268', 'station_name': 'unknown'}], ['新站'], '換電')
        self.assertTrue(all(not r['selected'] for r in out))

    def test_migrate_screenshot_old_draft_and_only_once(self):
        state = {'p::text': '電池 2.0E 501501011 1590039 車輛電池\n臺東火車站',
                 'p::draft': [dict(station_name='', equipment='501501011', kind='車輛電池', selected=True)]}
        prepare_vehicle_draft(state, 'p', ['臺東火車站'])
        self.assertEqual(len(state['p::draft']), 1)
        self.assertEqual(state['p::draft'][0]['equipment'], '1590039')
        self.assertEqual(state['p::draft'][0]['station_name'], '')
        self.assertFalse(state['p::draft'][0]['selected'])
        request = state['p::lookup_request']
        prepare_vehicle_draft(state, 'p', ['臺東火車站'])
        self.assertEqual(state['p::lookup_request'], request)
        self.assertEqual(state['p::draft_schema'], DRAFT_SCHEMA)

    def test_header_is_not_a_work_order(self):
        rows = parse_text('維護項目 車種 場站 車輛號碼 故障資訊 開單來源\n電池 2.0E 501501011 1590039 車輛電池', [])
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]['equipment'], '1590039')

    def test_lookup_cannot_accept_station_number_as_vehicle(self):
        rows = apply_lookup_to_draft([], [dict(station_name='A', bike_no='501501011')], ['A'], '換電')
        self.assertEqual(rows, [])

    def test_requery_clears_stale_location_and_keeps_successful_match(self):
        draft = [dict(equipment='1560268', station_name='OLD', selected=True, location_confirmed=True),
                 dict(equipment='1590102', station_name='OLD', selected=True, location_confirmed=True)]
        out = apply_lookup_to_draft(draft, [dict(bike_no='1560268', station_name='NEW')], ['NEW'], '換電', ['1560268', '1590102'])
        self.assertEqual(out[0]['station_name'], 'NEW')
        self.assertTrue(out[0]['location_confirmed'])
        self.assertEqual(out[1]['station_name'], '')
        self.assertFalse(out[1]['location_confirmed'])
        self.assertFalse(out[1]['selected'])


if __name__ == '__main__':
    unittest.main()
