import PyInstaller.__main__
import sys
import os

# 현재 디렉토리로 변경
os.chdir(r'c:\Users\ADMIN\Desktop\new')

# PyInstaller 실행
PyInstaller.__main__.run([
    '--name=날씨조회',
    '--onefile',
    '--windowed',
    '--noupx',
    '날씨조회_test1.py',
    '--clean'
])

print("\n빌드 완료! dist 폴더를 확인하세요.")










