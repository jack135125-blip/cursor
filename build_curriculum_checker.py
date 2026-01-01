# -*- coding: utf-8 -*-
import PyInstaller.__main__
import os

# 현재 디렉토리로 변경
os.chdir(os.path.dirname(os.path.abspath(__file__)))

# 파일명 (한글 포함)
script_file = 'curriculum_checker_12.26_테스트용 수정(1.2) copy.py'

# PyInstaller 실행
PyInstaller.__main__.run([
    '--name=curriculum_checker',
    '--onefile',
    '--windowed',
    '--icon=curriculum_checker_icon.ico',
    script_file,
    '--clean'
])

print("\n빌드 완료! dist 폴더를 확인하세요.")

