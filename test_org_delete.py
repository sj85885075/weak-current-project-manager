import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
import org_deliver


class OrgDeleteTests(unittest.TestCase):
    def test_delete_clear_reload_and_project_isolation(self):
        with tempfile.TemporaryDirectory() as folder, patch.object(org_deliver, 'ROOT', folder):
            Path(folder, '泰坦IDC').mkdir()
            for proj in ('bg', 'titan'):
                org_deliver.save_org(proj, {'people': [
                    {'pid':'P1', 'name':'甲', 'status':'在职'},
                    {'pid':'P2', 'name':'乙', 'status':'停用'}]})
            org=org_deliver.load_org('bg')
            org['people'][0]['status']='已删除'
            org_deliver.save_org('bg', org)
            self.assertEqual(org_deliver.org_stats(org_deliver.load_org('bg'))['total'], 1)
            for person in org['people']:person['status']='已删除'
            org_deliver.save_org('bg', org)
            after=org_deliver.load_org('bg')
            self.assertEqual(org_deliver.org_stats(after)['total'], 0)
            self.assertEqual(len(after['people']), 2)
            self.assertEqual(after['groups'], org['groups'])
            self.assertEqual(after['jobs'], org['jobs'])
            self.assertEqual(org_deliver.org_stats(org_deliver.load_org('titan'))['total'], 2)
            self.assertEqual(org_deliver.sync_org_from_people('bg', [['2026-09-18','甲','施工人员','内部','班组','','','',8]]), 0)
            self.assertEqual(org_deliver.org_stats(org_deliver.load_org('bg'))['total'], 0)


if __name__ == '__main__':unittest.main()
