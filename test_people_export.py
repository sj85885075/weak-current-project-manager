import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import openpyxl
import org_deliver as delivery


class PeopleExportTest(unittest.TestCase):
    def test_daily_template_and_grouping(self):
        data = {'people': [
            {'date': '2026-10-06', 'name': '甲', 'note': '原备注'},
            {'date': '2026-10-05', 'name': '乙'},
            {'date': '2026-10-06', 'name': '丙'},
            {'date': '2026-10-06', 'name': ''},
        ]}
        with patch.object(delivery._server(), 'bootstrap', return_value=data):
            title, headers, rows = delivery.collect_table('titan', 'people')
        self.assertEqual(headers[-2:], ['备注', '每日到场人数统计'])
        self.assertEqual([r[1] for r in rows], ['乙', '甲', '丙', ''])
        self.assertEqual([r[-1] for r in rows], [1, 2, 2, 2])
        info = dict(proj_name='测试项目', no='PRJ-TITAN', version='V1.0',
                    date='2026-10-06', org='项目部', editor='')
        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp, 'people.xlsx')
            delivery._to_xlsx(path, title, headers, rows, info)
            book = openpyxl.load_workbook(path)
            sheet = book.active
            self.assertEqual(sheet['J5'].value, '原备注')
            self.assertEqual(sheet['K4'].value, '=COUNTA(B4:B4)')
            self.assertEqual(sheet['K5'].value, '=COUNTA(B5:B7)')
            self.assertIn('A5:A7', [str(r) for r in sheet.merged_cells.ranges])
            self.assertIn('K5:K7', [str(r) for r in sheet.merged_cells.ranges])
            self.assertEqual(sheet.print_title_rows, '$1:$3')
            self.assertEqual(sheet['K3'].fill.fgColor.rgb, sheet['J3'].fill.fgColor.rgb)
            book.close()


if __name__ == '__main__':
    unittest.main()
