import base64
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import openpyxl
import org_deliver
import server


PNG_1PX = base64.b64decode('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M/wHwAF/gL+X8WnAAAAAElFTkSuQmCC')


class OrgProfileTests(unittest.TestCase):
    def test_profile_fields_and_photo_roundtrip(self):
        with tempfile.TemporaryDirectory() as root, patch.object(org_deliver, 'ROOT', root):
            result = org_deliver.save_org('titan', {'people': [{
                'name':'张三','gender':'男','phone':'13800000000','trade':'光缆熔接工',
                'job':'技术员','id_no':'ID-001','entry_date':'2026-09-22',
                'bio':'持证上岗，负责熔接和复测','photo_data':'data:image/png;base64,'+base64.b64encode(PNG_1PX).decode()
            }]})
            person=result['people'][0]
            self.assertNotIn('photo_data', person)
            self.assertTrue(person['photo'].endswith('.png'))
            self.assertEqual(Path(root,'泰坦IDC','人员照片',person['photo']).read_bytes(),PNG_1PX)
            loaded=org_deliver.load_org('titan')['people'][0]
            for key in ('gender','phone','trade','job','id_no','entry_date','bio','photo'):
                self.assertEqual(loaded[key],person[key])
            book=openpyxl.load_workbook(Path(root,'泰坦IDC','26_项目人员组织架构.xlsx'))
            self.assertEqual([c.value for c in book['人员信息'][1]],org_deliver.ORG_HEADERS)
            book.close()

    def test_old_14_column_file_upgrades_without_losing_data(self):
        with tempfile.TemporaryDirectory() as root, patch.object(org_deliver, 'ROOT', root):
            path=Path(root,'泰坦IDC');path.mkdir()
            book=openpyxl.Workbook();sheet=book.active;sheet.title='人员信息'
            sheet.append(org_deliver.ORG_HEADERS[:14]);sheet.append(['T-P01','旧人员','电工','施工组','其他','139','','','在职',1,'原备注','','','titan'])
            for title,headers in [('分组定义',org_deliver.GROUP_HEADERS),('岗位定义',org_deliver.JOB_HEADERS),('角色定义',org_deliver.ROLE_HEADERS)]:
                ws=book.create_sheet(title);ws.append(headers);ws.append(['其他',1,''])
            book.save(path/'26_项目人员组织架构.xlsx');book.close()
            old=org_deliver.load_org('titan');self.assertEqual(old['people'][0]['name'],'旧人员');self.assertIsNone(old['people'][0]['gender'])
            org_deliver.save_org('titan',old)
            loaded=org_deliver.load_org('titan')['people'][0]
            self.assertEqual((loaded['name'],loaded['phone'],loaded['note']),('旧人员','139','原备注'))

    def test_import_template_and_aliases_include_profile_fields(self):
        body,_=server.people_import_template('org')
        book=openpyxl.load_workbook(__import__('io').BytesIO(body));headers=[c.value for c in book['导入数据'][1]];book.close()
        for heading in ('性别','工种','证件号','入场日期','个人简介'):
            self.assertIn(heading,headers)

    def test_invalid_or_large_photo_rejected(self):
        with tempfile.TemporaryDirectory() as root, patch.object(org_deliver, 'ROOT', root):
            with self.assertRaisesRegex(ValueError,'格式'):
                org_deliver.save_org('titan',{'people':[{'name':'张三','photo_data':'data:image/png;base64,'+base64.b64encode(b'bad').decode()}]})


if __name__=='__main__': unittest.main()
