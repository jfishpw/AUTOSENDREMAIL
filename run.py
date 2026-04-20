#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
定时邮件收发系统启动脚本
"""
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.main import main

if __name__ == '__main__':
    main()
