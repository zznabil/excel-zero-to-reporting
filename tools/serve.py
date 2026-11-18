"""Serve only the published course files on a fixed loopback origin."""
from __future__ import annotations

import argparse
import json
import os
import shutil
import socket
import sys
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path, PurePosixPath
from typing import ClassVar, cast
from urllib.parse import unquote, urlsplit

HOST = '127.0.0.1'
PORT = 8765
ENTRY = 'START-HERE.html'
MIME = {'.html': 'text/html', '.js': 'text/javascript', '.css': 'text/css',
        '.json': 'application/json', '.csv': 'text/csv'}


class CourseServer(ThreadingHTTPServer):
    allow_reuse_address: ClassVar[bool] = sys.platform != 'win32'

    def server_bind(self) -> None:
        if sys.platform == 'win32':
            self.socket.setsockopt(socket.SOL_SOCKET, socket.SO_EXCLUSIVEADDRUSE, 1)
        super().server_bind()

    def __init__(self, address: tuple[str, int], root: Path):
        if address[0] != HOST:
            raise ValueError('The course server must bind only to 127.0.0.1.')
        self.root: Path = root.resolve()
        decoded: object = json.loads((self.root / 'release-files.json').read_text(encoding='utf-8'))
        if not isinstance(decoded, list) or not all(isinstance(n, str) for n in decoded):
            raise ValueError('release-files.json must contain a list of public file names.')
        names = cast(list[str], decoded)
        self.allowed: frozenset[str] = frozenset(names)
        if len(names) != len(self.allowed) or ENTRY not in self.allowed:
            raise ValueError('The public file list is incomplete or contains duplicates.')
        for name in names:
            path = PurePosixPath(name)
            if not name or '\\' in name or path.is_absolute() or '..' in path.parts:
                raise ValueError('The public file list contains an unsafe path.')
            if any(p in {'.git', 'My-work', 'test-results', 'node_modules', 'dist'} for p in path.parts):
                raise ValueError('Private working folders cannot be served.')
            if not self.public_path(name):
                raise ValueError(f'Missing or linked public file: {name}. Extract the whole ZIP.')
        super().__init__(address, CourseHandler)

    def public_path(self, name: str) -> Path | None:
        path = (self.root / name).resolve()
        return path if path.is_file() and path.is_relative_to(self.root) and path.relative_to(self.root).as_posix() == name else None

    @property
    def course_url(self) -> str:
        return f'http://{HOST}:{self.server_port}/{ENTRY}'


class CourseHandler(BaseHTTPRequestHandler):
    timeout: ClassVar[float | None] = 10.0

    def do_GET(self) -> None:
        self.send_course(False)

    def do_HEAD(self) -> None:
        self.send_course(True)

    def do_POST(self) -> None:
        self.send_error(405, 'This course server is read-only.')

    def send_course(self, head: bool) -> None:
        server = cast(CourseServer, self.server)
        if self.headers.get('Host') != f'{HOST}:{server.server_port}':
            self.send_error(403, f'Use {server.course_url}')
            return
        try:
            target = urlsplit(self.path)
        except ValueError:
            self.send_error(400, 'Invalid request URL.')
            return
        if target.scheme or target.netloc:
            self.send_error(400, 'Proxy requests are not supported.')
            return
        if target.path == '/':
            self.send_response(302)
            self.send_header('Location', '/' + ENTRY)
            self.send_header('Content-Length', '0')
            self.end_headers()
            return
        try:
            name = unquote(target.path, errors='strict').removeprefix('/')
            path = server.public_path(name) if name in server.allowed else None
            if path is None:
                self.send_error(404, 'Not a public course file.')
                return
            source = path.open('rb')
            size = os.fstat(source.fileno()).st_size
        except (OSError, UnicodeError):
            self.send_error(404, 'The course file could not be read. Extract the whole ZIP.')
            return
        with source:
            self.send_response(200)
            self.send_header('Content-Type', MIME.get(path.suffix, 'text/plain') + '; charset=utf-8')
            self.send_header('Content-Length', str(size))
            self.send_header('Cache-Control', 'no-store')
            self.send_header('X-Content-Type-Options', 'nosniff')
            self.end_headers()
            if not head:
                shutil.copyfileobj(source, self.wfile)


def main() -> int:
    parser = argparse.ArgumentParser(description='Open the offline course at a stable local origin. Requires Python 3.10+. Uses only the standard library; no administrator access or internet service is used.')
    parser.add_argument('--no-browser', action='store_true', help='serve without opening a browser; useful for automated checks')
    args = parser.parse_args()
    if sys.version_info < (3, 10):
        parser.error('Python 3.10 or later is required.')
    try:
        server = CourseServer((HOST, PORT), Path(__file__).resolve().parents[1])
    except (OSError, ValueError) as error:
        print(f'Cannot start the course: {error}\nKeep port {PORT} free; close a previous course console, not unrelated applications. Extract the complete ZIP.', file=sys.stderr)
        return 2
    with server:
        print('COURSE_READY ' + server.course_url, flush=True)
        print('Keep this console open while learning. Stop with Ctrl+C. Save JSON backups before changing browser or address.', flush=True)
        if not args.no_browser:
            try:
                if not webbrowser.open(server.course_url):
                    print('The browser did not open. Paste the COURSE_READY address into Chrome or Edge.', file=sys.stderr)
            except webbrowser.Error as error:
                print(f'Open the COURSE_READY address manually: {error}', file=sys.stderr)
        try:
            server.serve_forever()
        except KeyboardInterrupt:
            return 0
    return 0


if __name__ == '__main__':
    sys.exit(main())
