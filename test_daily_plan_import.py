import base64
import datetime
import io
import json
from pathlib import Path
import shutil
import tempfile
import unittest
from unittest.mock import patch
import openpyxl
import daily_plan_import as importer
import server


def parse(text):return importer.parse_upload('plan.csv',base64.b64encode(text.encode('utf-8-sig')).decode())


class ParserTests(unittest.TestCase):
    def test_zero_blank_and_sort(self):
        rows,skipped=parse('日期,每日计划量\n2026-09-20,0\n2026-09-19,2.5\n2026-09-21,\n')
        self.assertEqual([r['qty'] for r in rows],[2.5,0]);self.assertEqual(skipped,1)

    def test_reject_invalid_rows(self):
        for rows in ['2026-09-18,1\n2026-09-18,2','2026-02-30,1','2026-09-18,-2','2026-09-18,nan','2026-09-18,inf','2026-09-18,=1+1',',3']:
            with self.subTest(rows=rows),self.assertRaises(ValueError):parse('日期,每日计划量\n'+rows)

    def test_headers_and_empty(self):
        for text in ['日期,数量\n2026-09-18,2','日期,每日计划量\n','日期,每日计划量,每日计划量\n2026-09-18,1,1']:
            with self.assertRaises(ValueError):parse(text)

    def test_excel_date_and_formula(self):
        template=Path(__file__).parent/'outputs/daily-plan-import/每日计划数量导入模板.xlsx'
        book=openpyxl.load_workbook(template);sheet=book['每日计划导入']
        sheet['A2']=datetime.date(2026,9,18);sheet['B2']=4
        out=io.BytesIO();book.save(out)
        rows,_=importer.parse_upload('plan.xlsx',base64.b64encode(out.getvalue()).decode())
        self.assertEqual(rows[0]['date'],'2026-09-18');self.assertEqual(rows[0]['qty'],4)
        sheet['B2']='=1+1';out=io.BytesIO();book.save(out);book.close()
        with self.assertRaises(ValueError):importer.parse_upload('plan.xlsx',base64.b64encode(out.getvalue()).decode())


class IntegrationTests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory();self.real_dir=server.DIR
        for proj in ('bg','titan'):
            root=Path(self.real_dir)/('泰坦IDC' if proj=='titan' else '')
            target=Path(self.temp.name)/('泰坦IDC' if proj=='titan' else '');target.mkdir(exist_ok=True)
            for key in ('p30','p31','p32','p34'):
                file=Path(server._paths(proj)[key])
                if file.exists():shutil.copy2(file,target/file.name)
            cfg=root/'子系统计划设置.json'
            if cfg.exists():shutil.copy2(cfg,target/cfg.name)
        server.DIR=self.temp.name

    def tearDown(self):server.DIR=self.real_dir;self.temp.cleanup()

    def first(self,proj):
        sub=server.SUBS_OF[proj][0];schedule=server.sub_schedule(proj,sub)
        date=sorted(schedule)[0] if schedule else '2026-09-18'
        return sub,schedule,[{'row':2,'date':date,'qty':17 if proj=='titan' else 17.5}]

    def actual(self,proj):
        book=server.load_workbook(server._paths(proj)['p30'])
        try:return list(book['每日计量'].values) if '每日计量' in book.sheetnames else []
        finally:book.close()

    def assert_import(self,proj):
        sub,schedule,rows=self.first(proj);before_actual=self.actual(proj)
        other=server.SUBS_OF[proj][1];other_plan=server.sub_schedule(proj,other)
        check=importer.preview(server,proj,sub,rows)
        self.assertEqual(server.sub_schedule(proj,sub),schedule)
        result=importer.commit(server,proj,sub,rows,check['token'])
        self.assertTrue(result['ok']);expected={**schedule,rows[0]['date']:rows[0]['qty']}
        self.assertEqual(server.sub_schedule(proj,sub),expected)
        self.assertEqual(self.actual(proj),before_actual)
        self.assertEqual(server.sub_schedule(proj,other),other_plan)
        book=server.load_workbook(server._paths(proj)['p30']);sheet=book['工程量总览']
        col=5 if proj=='titan' else 4
        total=sum(float(sheet.cell(r,col).value or 0) for r in range(2,sheet.max_row+1) if server.canonical_sub(proj,sheet.cell(r,2).value)==sub)
        book.close();self.assertAlmostEqual(total,sum(expected.values()),places=2)

    def test_titan_settings_commit(self):self.assert_import('titan')
    def test_bg_settings_commit(self):self.assert_import('bg')

    def test_legacy_titan_and_bg(self):
        for proj in ('titan','bg'):
            path=Path(server._paths(proj)['p30']).parent/'子系统计划设置.json'
            if path.exists():path.unlink()
            self.assert_import(proj)

    def test_scope_integer_and_stale(self):
        sub,schedule,rows=self.first('titan')
        with self.assertRaisesRegex(ValueError,'整数'):importer.preview(server,'titan',sub,[{**rows[0],'qty':1.1}])
        with self.assertRaisesRegex(ValueError,'工期'):importer.preview(server,'titan',sub,[{**rows[0],'date':'2000-01-01'}])
        with self.assertRaisesRegex(ValueError,'无效'):importer.preview(server,'invalid',sub,rows)
        with self.assertRaisesRegex(ValueError,'已变化'):importer.commit(server,'titan',sub,rows,'old-token')

    def test_full_rollback(self):
        sub,schedule,rows=self.first('titan');check=importer.preview(server,'titan',sub,rows)
        root=Path(server._paths('titan')['p30']).parent
        files=[Path(server._paths('titan')[k]) for k in ('p30','p31','p32','p34')]+[root/'子系统计划设置.json']
        original={p:p.read_bytes() if p.exists() else None for p in files}
        def fail(data):
            files[1].write_bytes(b'partial update')
            raise RuntimeError('simulated sync failure')
        with patch.object(server,'save_sub_plan',side_effect=fail),self.assertRaises(RuntimeError):importer.commit(server,'titan',sub,rows,check['token'])
        for p,content in original.items():self.assertEqual(p.read_bytes() if p.exists() else None,content)


if __name__=='__main__':unittest.main()
