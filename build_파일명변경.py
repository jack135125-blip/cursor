# -*- coding: utf-8 -*-
"""파일명변경프로그램.py → exe 빌드 스크립트"""
import os

import PyInstaller.__main__

os.chdir(os.path.dirname(os.path.abspath(__file__)))

PyInstaller.__main__.run([
    "--name=파일명변경",
    "--onefile",
    "--windowed",
    "--noupx",
    "파일명변경프로그램.py",
    "--clean",
])

print("\n빌드 완료! dist\\파일명변경.exe 를 확인하세요.")
