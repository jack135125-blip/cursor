# -*- coding: utf-8 -*-
"""교육과정 편성표 점검 프로그램.py → exe 빌드 스크립트"""
import os

import PyInstaller.__main__

os.chdir(os.path.dirname(os.path.abspath(__file__)))

script_file = "교육과정 편성표 점검 프로그램.py"
icon_file = "curriculum_checker_icon.ico"

args = [
    "--name=교육과정 편성표 점검 프로그램",
    "--onefile",
    "--windowed",
    "--noupx",
    "--collect-all=openpyxl",
    "--hidden-import=requests",
    "--hidden-import=urllib3",
    script_file,
    "--clean",
]

if os.path.isfile(icon_file):
    args.insert(3, f"--icon={icon_file}")

PyInstaller.__main__.run(args)

print("\n빌드 완료! dist\\교육과정 편성표 점검 프로그램.exe 를 확인하세요.")
