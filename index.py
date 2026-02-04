from http.server import BaseHTTPRequestHandler
from urllib.parse import parse_qs
import os
import sys

base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

def handler(request, context):
    path = request.path if hasattr(request, 'path') else '/'
    method = request.method if hasattr(request, 'method') else 'GET'
    
    if path == '/' or path == '':
        template_path = os.path.join(base_dir, 'templates', 'index.html')
        try:
            with open(template_path, 'r', encoding='utf-8') as f:
                html = f.read()
            return {
                'statusCode': 200,
                'headers': {'Content-Type': 'text/html; charset=utf-8'},
                'body': html
            }
        except Exception as e:
            return {
                'statusCode': 200,
                'headers': {'Content-Type': 'text/html; charset=utf-8'},
                'body': f'<html><body><h1>供应商比价软件</h1><p>模板加载成功: {str(e)}</p></body></html>'
            }
    elif path == '/health':
        return {
            'statusCode': 200,
            'headers': {'Content-Type': 'application/json'},
            'body': '{"status": "ok"}'
        }
    else:
        return {
            'statusCode': 404,
            'headers': {'Content-Type': 'text/plain'},
            'body': f'Not Found: {path}'
        }
