#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Server xuất báo giá - HƯNG THỊNH SMART
Chạy local:  python3 server.py  -> http://localhost:5000
Chạy online (Render...): dùng biến môi trường PORT, có lớp mật khẩu bảo vệ toàn site.
"""

from flask import Flask, request, jsonify, send_file
from pathlib import Path
import json
import subprocess
import sys
import os

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = Path(__file__).parent / 'output'
app.config['UPLOAD_FOLDER'].mkdir(exist_ok=True)


@app.route('/api/export-quote', methods=['POST'])
def export_quote():
    """Endpoint để xuất báo giá"""
    try:
        data = request.get_json()
        json_str = json.dumps(data, ensure_ascii=False)

        result = subprocess.run(
            [sys.executable, 'scripts/export_quote_to_excel.py'],
            input=json_str.encode('utf-8'),
            capture_output=True,
            timeout=30,
            cwd=Path(__file__).parent
        )

        if result.returncode != 0:
            return jsonify({'success': False, 'error': result.stderr.decode('utf-8')}), 500

        output = result.stdout.decode('utf-8')

        lines = output.split('\n')
        excel_file = None
        for line in lines:
            if 'Lưu Excel:' in line or 'OK Excel:' in line:
                excel_file = line.split(':', 1)[1].strip()
                break

        if excel_file:
            return jsonify({
                'success': True,
                'message': output,
                'file': Path(excel_file).name
            })
        else:
            return jsonify({'success': False, 'error': 'Không tìm thấy file Excel', 'log': output}), 500

    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/download/<filename>', methods=['GET'])
def download_file(filename):
    try:
        file_path = Path(__file__).parent / 'scripts' / filename
        if not file_path.exists():
            file_path = Path(__file__).parent / 'output' / filename
        if not file_path.exists():
            return jsonify({'error': 'File not found'}), 404
        return send_file(file_path, as_attachment=True)
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/', methods=['GET'])
def index():
    app_file = Path(__file__).parent / 'app' / 'app_quote_embedded.html'
    if app_file.exists():
        return app_file.read_text(encoding='utf-8')
    return 'App not found', 404


@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok'})


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    print("\n" + "=" * 70)
    print("🚀 SERVER XUẤT BÁO GIÁ - HƯNG THỊNH SMART")
    print("=" * 70)
    print(f"\n✅ Mở trình duyệt: http://localhost:{port}")
    print("\n⚠️  Nhấn Ctrl+C để dừng server\n")
    print("=" * 70 + "\n")

    app.run(debug=False, host='0.0.0.0', port=port)
