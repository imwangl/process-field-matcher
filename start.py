#!/usr/bin/env python3
import subprocess
import sys
import os

# 安装依赖
subprocess.run([sys.executable, '-m', 'pip', 'install', 'flask', 'pandas', 'openpyxl', 'python-Levenshtein', '-q'], capture_output=True)

# 启动服务
os.chdir(os.path.expanduser('~/.openclaw/workspace/加工字段匹配工具/'))
os.execv(sys.executable, [sys.executable, 'app.py'])