#!/usr/bin/env python3
"""
Presentation Server — 簡報編輯伺服器
提供 HTML 簡報瀏覽 + 即時編輯 + 儲存回 MD

用法:
    python presentation_server.py path/to/presentation.md [--port 8080]

啟動後自動開啟瀏覽器，編輯後按「儲存」會：
1. 將編輯內容反向轉回 Presentation MD
2. 寫入 .md 檔案
3. 重新生成 HTML
4. 自動重新載入頁面
"""

import argparse
import base64
import io
import json
import re
import shutil
import subprocess
import sys
import threading
import urllib.parse
import webbrowser
from datetime import datetime
from http.server import HTTPServer, SimpleHTTPRequestHandler
from pathlib import Path

if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

SKILL_DIR = Path(__file__).parent.parent
GENERATE_SCRIPT = SKILL_DIR / 'scripts' / 'generate_html.py'


class PresentationHandler(SimpleHTTPRequestHandler):
    """處理簡報的 GET（瀏覽）和 POST（儲存）請求"""

    md_path: Path = None
    html_path: Path = None
    serve_dir: Path = None

    def do_GET(self):
        if self.path == '/' or self.path == '/index.html':
            self._serve_html()
        else:
            # 讓其他檔案（圖片等）能被載入
            self.directory = str(self.serve_dir)
            super().do_GET()

    def do_POST(self):
        if self.path == '/api/save':
            self._handle_save()
        elif self.path == '/api/save-image':
            self._handle_save_image()
        elif self.path == '/api/save-video':
            self._handle_save_video()
        elif self.path == '/api/snapshot':
            self._handle_snapshot()
        else:
            self.send_error(404)

    def _serve_html(self):
        try:
            data = self.html_path.read_bytes()
            self.send_response(200)
            self.send_header('Content-Type', 'text/html; charset=utf-8')
            self.send_header('Content-Length', len(data))
            self.send_header('Cache-Control', 'no-cache')
            self.end_headers()
            self.wfile.write(data)
        except FileNotFoundError:
            self.send_error(404, 'HTML file not found — run generate_html.py first')

    def _handle_save(self):
        try:
            length = int(self.headers.get('Content-Length', 0))
            body = self.rfile.read(length)
            data = json.loads(body.decode('utf-8'))
            md_content = data.get('markdown', '')

            if not md_content.strip():
                self._json_response(400, {'ok': False, 'error': 'MD 內容為空'})
                return

            # 寫入 MD
            self.md_path.write_text(md_content, encoding='utf-8')

            # 重新生成 HTML
            result = subprocess.run(
                [sys.executable, str(GENERATE_SCRIPT),
                 str(self.md_path), '--output', str(self.html_path)],
                capture_output=True, text=True, encoding='utf-8', errors='replace'
            )

            if result.returncode == 0:
                self._json_response(200, {
                    'ok': True,
                    'message': f'MD 已儲存，HTML 已重新生成',
                    'reload': True,
                })
                self.log_message('Saved MD + regenerated HTML')
            else:
                # MD 已儲存，但 HTML 生成失敗
                self._json_response(200, {
                    'ok': True,
                    'message': f'MD 已儲存，但 HTML 重新生成失敗：{result.stderr[:200]}',
                    'reload': False,
                })
                self.log_message(f'HTML generation failed: {result.stderr[:200]}')

        except json.JSONDecodeError:
            self._json_response(400, {'ok': False, 'error': '無效的 JSON'})
        except Exception as e:
            self._json_response(500, {'ok': False, 'error': str(e)})

    def _handle_save_video(self):
        """接收影片二進位（raw bytes），儲存到 images/ 子目錄，回傳相對路徑"""
        try:
            length = int(self.headers.get('Content-Length', 0))
            raw = self.rfile.read(length)

            filename_header = self.headers.get('X-Filename', 'video.mp4')
            try:
                filename_header = urllib.parse.unquote(filename_header)
            except Exception:
                pass

            ext = Path(filename_header).suffix.lower()
            allowed = {'.mp4', '.webm', '.mov', '.avi', '.mkv', '.m4v'}
            if ext not in allowed:
                ext = '.mp4'

            img_dir = self.md_path.parent / 'images'
            img_dir.mkdir(exist_ok=True)

            stem = re.sub(r'[^\w\-]', '_', Path(filename_header).stem)[:30]
            ts = datetime.now().strftime('%H%M%S')
            filename = f'{stem}_{ts}{ext}'
            dest = img_dir / filename
            dest.write_bytes(raw)

            rel_path = f'images/{filename}'
            self._json_response(200, {'ok': True, 'path': rel_path})
            self.log_message(f'Saved video: {rel_path}')

        except Exception as e:
            self._json_response(500, {'ok': False, 'error': str(e)})

    def _handle_snapshot(self):
        """將當前工作版本複製到 snapshots/{name}/ 資料夾"""
        try:
            length = int(self.headers.get('Content-Length', 0))
            body = self.rfile.read(length)
            data = json.loads(body.decode('utf-8'))
            name = data.get('name', '').strip()

            if not name:
                self._json_response(400, {'ok': False, 'error': '快照名稱不能為空'})
                return

            # 去除不安全字元（保留中文、英數、底線、連字號）
            safe_name = re.sub(r'[\\/:*?"<>|]', '_', name)
            base_dir = self.md_path.parent
            snap_dir = base_dir / 'snapshots' / safe_name

            if snap_dir.exists():
                self._json_response(400, {
                    'ok': False,
                    'error': f'快照「{safe_name}」已存在，請使用不同名稱'
                })
                return

            snap_dir.mkdir(parents=True)

            # 複製 MD
            shutil.copy2(self.md_path, snap_dir / self.md_path.name)

            # 複製 HTML（重新生成，讓路徑指向 snapshot 內的 images/）
            snap_html = snap_dir / self.html_path.name
            result = subprocess.run(
                [sys.executable, str(GENERATE_SCRIPT),
                 str(snap_dir / self.md_path.name),
                 '--output', str(snap_html)],
                capture_output=True, text=True, encoding='utf-8', errors='replace'
            )
            if result.returncode != 0:
                # 退而求其次：直接複製 HTML
                shutil.copy2(self.html_path, snap_html)

            # 複製 images/ 資料夾
            img_src = base_dir / 'images'
            if img_src.exists():
                shutil.copytree(img_src, snap_dir / 'images')

            self._json_response(200, {
                'ok': True,
                'path': f'snapshots/{safe_name}',
            })
            self.log_message(f'Snapshot saved: snapshots/{safe_name}')

        except Exception as e:
            self._json_response(500, {'ok': False, 'error': str(e)})

    def _handle_save_image(self):
        """接收 base64 圖片，儲存到 images/ 子目錄，回傳相對路徑"""
        try:
            length = int(self.headers.get('Content-Length', 0))
            body = self.rfile.read(length)
            data = json.loads(body.decode('utf-8'))

            b64 = data.get('base64', '')
            hint = data.get('filename', 'image.png')

            # 解析 data URL（data:image/png;base64,...）
            match = re.match(r'data:image/(\w+);base64,(.*)', b64, re.DOTALL)
            if not match:
                self._json_response(400, {'ok': False, 'error': '無效的圖片資料'})
                return

            ext = match.group(1).lower()
            if ext == 'jpeg': ext = 'jpg'
            raw = base64.b64decode(match.group(2))

            # 建立 images/ 目錄
            img_dir = self.md_path.parent / 'images'
            img_dir.mkdir(exist_ok=True)

            # 使用原始檔名（去除非安全字元），加時間戳避免衝突
            stem = re.sub(r'[^\w\-]', '_', Path(hint).stem)[:30]
            ts = datetime.now().strftime('%H%M%S')
            filename = f'{stem}_{ts}.{ext}'
            dest = img_dir / filename

            dest.write_bytes(raw)

            rel_path = f'images/{filename}'
            self._json_response(200, {'ok': True, 'path': rel_path})
            self.log_message(f'Saved image: {rel_path}')

        except Exception as e:
            self._json_response(500, {'ok': False, 'error': str(e)})

    def _json_response(self, code, obj):
        data = json.dumps(obj, ensure_ascii=False).encode('utf-8')
        self.send_response(code)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Content-Length', len(data))
        self.end_headers()
        self.wfile.write(data)

    def log_message(self, fmt, *args):
        # 簡化日誌格式
        msg = fmt % args if args else fmt
        print(f'  [{self.log_date_time_string()}] {msg}')


def main():
    parser = argparse.ArgumentParser(description='Presentation Editor Server')
    parser.add_argument('input', help='Presentation MD file path')
    parser.add_argument('--port', '-p', type=int, default=8080, help='Server port (default: 8080)')
    parser.add_argument('--no-open', action='store_true', help='Do not open browser automatically')
    args = parser.parse_args()

    md_path = Path(args.input).resolve()
    if not md_path.exists():
        print(f'Error: {md_path} not found', file=sys.stderr)
        sys.exit(1)

    html_path = md_path.with_suffix('.html')

    # 先生成一次 HTML
    print(f'Generating HTML from {md_path.name}...')
    result = subprocess.run(
        [sys.executable, str(GENERATE_SCRIPT), str(md_path), '--output', str(html_path)],
        capture_output=True, text=True, encoding='utf-8', errors='replace'
    )
    if result.returncode != 0:
        print(f'Error generating HTML: {result.stderr}', file=sys.stderr)
        sys.exit(1)
    print(result.stdout.strip())

    # 設定 handler
    PresentationHandler.md_path = md_path
    PresentationHandler.html_path = html_path
    PresentationHandler.serve_dir = md_path.parent

    server = HTTPServer(('127.0.0.1', args.port), PresentationHandler)
    url = f'http://localhost:{args.port}'

    print(f'\n  Presentation Server')
    print(f'  URL:  {url}')
    print(f'  MD:   {md_path}')
    print(f'  HTML: {html_path}')
    print(f'\n  Press E in browser to edit | Save writes back to MD')
    print(f'  Press Ctrl+C to stop\n')

    if not args.no_open:
        threading.Timer(0.5, lambda: webbrowser.open(url)).start()

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print('\n  Server stopped.')
        server.server_close()


if __name__ == '__main__':
    main()
