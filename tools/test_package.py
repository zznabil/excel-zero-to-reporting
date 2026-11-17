import copy
import json
from pathlib import Path
import tempfile
import unittest
import zipfile
from tools import package

class PackagingChecks(unittest.TestCase):
    def test_reproducible_real_archive(self):
        with tempfile.TemporaryDirectory(prefix='excel release proof ') as folder:
            a,_=package.build(Path(folder)/'first.zip')
            b,_=package.build(Path(folder)/'second.zip')
            self.assertEqual(a.read_bytes(),b.read_bytes())
            self.assertGreater(package.verify(a),20)
            self.assertEqual(a.with_suffix('.zip.sha256').read_text().split()[0],package.digest(a.read_bytes()))

    def test_actual_verifier_rejects_omission_and_corruption(self):
        with tempfile.TemporaryDirectory() as folder:
            good,_=package.build(Path(folder)/'good.zip')
            with zipfile.ZipFile(good) as archive:
                members={n:archive.read(n) for n in archive.namelist()}
            self.assertGreater(package.validate_members(members),20)
            for kind in ('omitted','corrupt'):
                broken=copy.copy(members)
                if kind=='omitted':del broken['demos.js']
                else:broken['demos.js']+=b'// altered release payload'
                file=Path(folder)/(kind+'.zip')
                with zipfile.ZipFile(file,'w') as archive:
                    for name,data in broken.items():archive.writestr(name,data)
                with self.assertRaises((ValueError,KeyError)):package.verify(file)

    def test_sensitive_path_and_credential_controls_have_positive_controls(self):
        package.inspect_payload('safe.csv',b'ID,Amount\nA001,50\n')
        private=bytes([67,58,92])+b'Users'+bytes([92])+b'SENTINEL'+bytes([92])+b'private.txt'
        token=('gh'+'p_'+'A'*40).encode()
        for payload in (private,token):
            with self.assertRaises(ValueError):package.inspect_payload('data.txt',payload)
        for name in ('../escape.txt','nested.mp4','nested.xlsx','nested.zip'):
            with self.assertRaises(ValueError):package.inspect_payload(name,b'forbidden fixture')

    def test_duplicate_zip_entry_is_not_silently_retained(self):
        with tempfile.TemporaryDirectory() as folder:
            file=Path(folder)/'duplicate.zip'
            with zipfile.ZipFile(file,'w') as archive:
                archive.writestr('START-HERE.html',b'one')
                archive.writestr('START-HERE.html',b'two')
            with self.assertRaisesRegex(ValueError,'Duplicate ZIP entry'):package.verify(file)

    def test_linked_required_script_cannot_be_removed_from_the_allowlist(self):
        with tempfile.TemporaryDirectory() as folder:
            good,_=package.build(Path(folder)/'good.zip')
            with zipfile.ZipFile(good) as archive:
                members={name:archive.read(name) for name in archive.namelist()}
            self.assertGreater(package.validate_members(members),20)
            del members['checks.js']
            names=json.loads(members['release-files.json'])
            names.remove('checks.js')
            members['release-files.json']=json.dumps(names).encode()
            inventory=json.loads(members['INVENTORY.json'])
            inventory['files']=[entry for entry in inventory['files'] if entry['path']!='checks.js']
            for entry in inventory['files']:
                if entry['path']=='release-files.json':
                    entry['bytes']=len(members['release-files.json'])
                    entry['sha256']=package.digest(members['release-files.json'])
            members['INVENTORY.json']=json.dumps(inventory).encode()
            with self.assertRaisesRegex(ValueError,'Missing linked asset: checks.js'):
                package.validate_members(members)

if __name__=='__main__':
    unittest.main(verbosity=2)
