"""Build and verify an explicit original-file allowlist. Python standard library only."""
import argparse
import hashlib
import json
import re
import zipfile
from html.parser import HTMLParser
from pathlib import Path, PurePosixPath
from urllib.parse import unquote, urlsplit

ROOT = Path(__file__).resolve().parent.parent
FORBIDDEN_SUFFIXES = {'.xlsx', '.xlsm', '.xls', '.mp4', '.srt', '.pdf', '.zip', '.log'}

def digest(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()

def inspect_payload(name: str, data: bytes) -> None:
    path = PurePosixPath(name)
    if path.is_absolute() or '..' in path.parts or '\\' in name or ':' in name:
        raise ValueError('Unsafe archive path: ' + name)
    if path.suffix.lower() in FORBIDDEN_SUFFIXES:
        raise ValueError('Non-allowlisted media/workbook/archive type: ' + name)
    if re.search(rb'[A-Za-z]:[\\/]+Users[\\/]+[^\s\\/]+', data, re.IGNORECASE):
        raise ValueError('Private absolute user path in ' + name)
    if re.search(rb'gh[pousr]_[A-Za-z0-9]{30,}|github_pat_[A-Za-z0-9_]{30,}', data):
        raise ValueError('Possible credential in ' + name)

class AssetLinks(HTMLParser):
    def __init__(self, source: str, members: dict[str, bytes]):
        super().__init__()
        self.source = source
        self.members = members

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag not in {'a', 'script', 'link', 'img', 'iframe'}:
            return
        attribute = 'src' if tag in {'script', 'img', 'iframe'} else 'href'
        reference = dict(attrs).get(attribute)
        if not reference:
            return
        url = urlsplit(reference)
        if url.scheme or url.netloc:
            if tag != 'a':
                raise ValueError('Required asset depends on a remote service: ' + reference)
            return
        if not url.path:
            return
        target = PurePosixPath(self.source).parent / unquote(url.path)
        if target.is_absolute() or '..' in target.parts or target.as_posix() not in self.members:
            raise ValueError('Missing linked asset: ' + reference)

def validate_members(members: dict[str, bytes]) -> int:
    if set(members) != set(json.loads(members['release-files.json'])) | {'INVENTORY.json'}:
        raise ValueError('Archive differs from its explicit allowlist')
    inventory = json.loads(members['INVENTORY.json'])
    if inventory.get('format') != 'excel-zero-to-reporting-inventory-v1':
        raise ValueError('Unsupported inventory')
    entries = inventory.get('files', [])
    if len(entries) != len(members)-1 or len({e['path'] for e in entries}) != len(entries):
        raise ValueError('Inventory does not cover archive members uniquely')
    for entry in entries:
        name = entry['path']
        data = members[name]
        inspect_payload(name, data)
        if len(data) != entry['bytes'] or digest(data) != entry['sha256']:
            raise ValueError('Integrity mismatch: ' + name)
    if {e['path'] for e in entries} != set(members)-{'INVENTORY.json'}:
        raise ValueError('Missing inventory entry')
    for name, data in members.items():
        if name.endswith('.html'):
            parser = AssetLinks(name, members)
            parser.feed(data.decode('utf-8'))
            parser.close()
    return len(entries)

def build(destination: Path | None = None) -> tuple[Path, int]:
    names = json.loads((ROOT/'release-files.json').read_text(encoding='utf-8'))
    if len(names) != len(set(names)):
        raise ValueError('Duplicate allowlist entry')
    members = {}
    for name in sorted(names):
        file = ROOT/name
        if file.is_symlink() or not file.is_file() or ROOT not in file.resolve().parents:
            raise ValueError('Missing or escaping input: ' + name)
        data = file.read_bytes()
        inspect_payload(name, data)
        members[name] = data
    inventory = {'format':'excel-zero-to-reporting-inventory-v1',
                 'provenance':'Original course software, prose, synthetic data and measured native test oracle. Third-party packages/media are not bundled.',
                 'files':[{'path':name,'bytes':len(data),'sha256':digest(data)} for name,data in members.items()]}
    members['INVENTORY.json'] = (json.dumps(inventory,indent=2,ensure_ascii=False)+'\n').encode('utf-8')
    validate_members(members)
    version = json.loads(members['package.json'])['version']
    if not re.fullmatch(r'\d+\.\d+\.\d+',version):
        raise ValueError('Unsupported release version')
    output = Path(destination) if destination else ROOT/'dist'/('excel-zero-to-reporting-'+version+'.zip')
    output.parent.mkdir(parents=True,exist_ok=True)
    with zipfile.ZipFile(output,'w',compression=zipfile.ZIP_STORED) as archive:
        for name in sorted(members):
            info = zipfile.ZipInfo(name,date_time=(1980,1,1,0,0,0))
            info.create_system = 3
            info.external_attr = 0o100644 << 16
            archive.writestr(info,members[name])
    count = verify(output)
    checksum = output.with_suffix('.zip.sha256')
    checksum.write_text(digest(output.read_bytes())+'  '+output.name+'\n',encoding='ascii')
    return output,count

def verify(file: Path | str) -> int:
    with zipfile.ZipFile(file) as archive:
        names = archive.namelist()
        if len(names) != len(set(names)):
            raise ValueError('Duplicate ZIP entry')
        if len(names) > 200 or sum(info.file_size for info in archive.infolist()) > 32*1024*1024:
            raise ValueError('Archive exceeds the bounded course size')
        members = {name:archive.read(name) for name in names}
    return validate_members(members)

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--verify',type=Path,help='Verify a downloaded ZIP against its embedded inventory and allowlist; also compare its external SHA256 before trusting its origin.')
    parser.add_argument('--output',type=Path,help='Build to this path instead of dist.')
    args = parser.parse_args()
    if args.verify:
        print('RELEASE_ARCHIVE_VERIFIED:',verify(args.verify),'original files')
    else:
        output,count = build(args.output)
        print('RELEASE_PACKAGE_VERIFIED:',output.name,count,'original files',digest(output.read_bytes()))
