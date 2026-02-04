import sys
import os

base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, base_dir)
os.chdir(base_dir)

from web_app import app

from flask import request

def application(environ, start_response):
    return app.wsgi_app(environ, start_response)
