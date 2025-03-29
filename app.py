#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import sys

# 设置项目根目录
ROOT_DIR = os.path.abspath(os.path.dirname(__file__))
sys.path.insert(0, ROOT_DIR)

from src.web import app

if __name__ == '__main__':
    # 在生产环境中应使用WSGI服务器（如gunicorn或uwsgi）
    # 这里仅为开发环境使用
    app.run(host='0.0.0.0', port=5000, debug=True)
