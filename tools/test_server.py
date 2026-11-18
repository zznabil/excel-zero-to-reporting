import csv
import http.client
import io
import json
import tempfile
import threading
import unittest
from collections.abc import Iterator
from contextlib import contextmanager
from pathlib import Path

from .serve import CourseServer


@contextmanager
def served_course(port: int = 0) -> Iterator[tuple[Path, CourseServer]]:
    with tempfile.TemporaryDirectory(prefix='excel server boundary ') as folder:
        root = Path(folder)
        (root / 'Data').mkdir()
        source = Path(__file__).resolve().parents[1] / 'Data' / 'Office-requests.csv'
        (root / 'Data/Office-requests.csv').write_bytes(source.read_bytes())
        (root / 'START-HERE.html').write_text('<h1>Course</h1>', encoding='utf-8')
        (root / 'My-work').mkdir()
        (root / 'My-work/private.csv').write_text('PRIVATE-WORKBOOK-CANARY', encoding='utf-8')
        (root / 'release-files.json').write_text(json.dumps(['START-HERE.html', 'Data/Office-requests.csv']), encoding='utf-8')
        with CourseServer(('127.0.0.1', port), root) as server:
            thread = threading.Thread(target=server.serve_forever, daemon=True)
            thread.start()
            try:
                yield root, server
            finally:
                server.shutdown()
                thread.join()


def request(server: CourseServer, method: str, path: str, headers: dict[str, str] | None = None) -> tuple[int, dict[str, str], bytes]:
    connection = http.client.HTTPConnection('127.0.0.1', server.server_port, timeout=5)
    try:
        connection.request(method, path, headers=headers or {})
        response = connection.getresponse()
        return response.status, dict(response.getheaders()), response.read()
    finally:
        connection.close()


class ServerBoundaryTests(unittest.TestCase):
    def test_public_csv_loads_and_reconciles_with_head_contract(self):
        with served_course() as (_, server):
            status, headers, body = request(server, 'GET', '/Data/Office-requests.csv')
            self.assertEqual(status, 200)
            rows = list(csv.DictReader(io.StringIO(body.decode('utf-8'))))
            self.assertEqual(len(rows), 8)
            self.assertEqual(sum(int(row['Amount']) for row in rows), 1170)
            self.assertEqual(headers['Content-Type'], 'text/csv; charset=utf-8')
            status, head, content = request(server, 'HEAD', '/Data/Office-requests.csv')
            self.assertEqual((status, int(head['Content-Length']), content), (200, len(body), b''))

    def test_private_encoded_paths_foreign_hosts_and_writes_are_denied(self):
        with served_course() as (root, server):
            self.assertEqual(request(server, 'GET', '/START-HERE.html')[0], 200)
            for path in ['/My-work/private.csv', '/%4dy-work/private.csv', '/Data/%2e%2e/My-work/private.csv', '/.git/config']:
                with self.subTest(path=path):
                    status, _, body = request(server, 'GET', path)
                    self.assertEqual(status, 404)
                    self.assertNotIn(b'PRIVATE-WORKBOOK-CANARY', body)
            self.assertEqual(request(server, 'GET', '/START-HERE.html', {'Host': 'untrusted.example'})[0], 403)
            self.assertEqual(request(server, 'POST', '/START-HERE.html')[0], 405)
            self.assertEqual((root / 'My-work/private.csv').read_text(), 'PRIVATE-WORKBOOK-CANARY')

    def test_second_listener_cannot_take_over_the_course_origin(self):
        with served_course() as (_, server):
            duplicate = None
            try:
                with self.assertRaises(OSError):
                    duplicate = CourseServer(('127.0.0.1', server.server_port), server.root)
            finally:
                if duplicate is not None:
                    duplicate.server_close()
            self.assertEqual(request(server, 'GET', '/Data/Office-requests.csv')[0], 200)

    def test_origin_can_restart_after_the_owned_server_closes(self):
        with served_course() as (_, server):
            port = server.server_port
            self.assertEqual(request(server, 'GET', '/Data/Office-requests.csv')[0], 200)
        with served_course(port) as (_, restarted):
            self.assertEqual(request(restarted, 'GET', '/Data/Office-requests.csv')[0], 200)


if __name__ == '__main__':
    unittest.main()
