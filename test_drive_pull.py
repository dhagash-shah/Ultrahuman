import unittest
from unittest.mock import Mock, patch
import drive_pull as drive

class DriveDiscoveryTests(unittest.TestCase):
 def test_reads_all_pages(self):
  svc=Mock()
  svc.files.return_value.list.return_value.execute.side_effect=[{'files':[{'id':'old'}],'nextPageToken':'second'},{'files':[{'id':'new'}]}]
  self.assertEqual([f['id'] for f in drive.list_files(svc,'folder')],['old','new'])
  self.assertEqual(svc.files.return_value.list.call_args.kwargs['pageToken'],'second')
  self.assertIn('nextPageToken',svc.files.return_value.list.call_args.kwargs['fields'])
 def test_modified_dates_remain_sortable(self):
  files=[{'id':'missing'},{'id':'good','modifiedTime':'2026-11-01T00:00:00Z'},{'id':'bad','modifiedTime':'bad'}]
  self.assertEqual(sorted(files,key=drive.modified_key,reverse=True)[0]['id'],'good')
 def test_month_selection(self):
  self.assertEqual(drive.month_key('MIS September 2026.xlsx'),(2026,9))

class CapFolderMigrationTests(unittest.TestCase):
 def test_retired_folder_is_replaced(self):
  for previous in ('', '1vn_8fIlOVM5EjtJtG5n1tMmnw2kDZLMz'):
   with patch.dict(drive.os.environ,{'CASHIFY_CAP_TABLES_FOLDER_ID':previous}):
    self.assertEqual(drive.configured_cap_folder(),'19-4RzHtmpfbnJcw1DjGR24PbIAY2mn9_')
 def test_custom_folder_is_preserved(self):
  with patch.dict(drive.os.environ,{'CASHIFY_CAP_TABLES_FOLDER_ID':'custom-folder'}):
   self.assertEqual(drive.configured_cap_folder(),'custom-folder')

if __name__=='__main__':unittest.main()
