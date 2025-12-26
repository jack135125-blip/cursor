# -*- coding: utf-8 -*-
"""
점검 확인 아이콘 생성 스크립트
"""
from PIL import Image, ImageDraw, ImageFont
import os

def create_check_icon():
    """점검 확인 느낌의 아이콘 생성"""
    # 아이콘 크기 (여러 크기 지원)
    sizes = [(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)]
    images = []
    
    for size in sizes:
        # 이미지 생성 (투명 배경)
        img = Image.new('RGBA', size, (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)
        
        # 배경 원 그리기 (파란색 계열)
        center = (size[0] // 2, size[1] // 2)
        radius = min(size) // 2 - 4
        
        # 배경 원 (밝은 파란색)
        draw.ellipse(
            [center[0] - radius, center[1] - radius, 
             center[0] + radius, center[1] + radius],
            fill=(52, 152, 219, 255),  # 파란색
            outline=(41, 128, 185, 255),  # 진한 파란색 테두리
            width=max(2, size[0] // 32)
        )
        
        # 체크마크 그리기 (흰색)
        check_size = radius * 0.6
        check_thickness = max(3, size[0] // 16)
        
        # 체크마크 좌표 계산
        check_start = (center[0] - check_size * 0.4, center[1])
        check_mid = (center[0] - check_size * 0.1, center[1] + check_size * 0.3)
        check_end = (center[0] + check_size * 0.5, center[1] - check_size * 0.3)
        
        # 체크마크 그리기 (두꺼운 선)
        draw.line([check_start, check_mid], fill=(255, 255, 255, 255), width=check_thickness)
        draw.line([check_mid, check_end], fill=(255, 255, 255, 255), width=check_thickness)
        
        images.append(img)
    
    # ICO 파일로 저장
    ico_path = 'curriculum_checker_icon.ico'
    images[0].save(
        ico_path,
        format='ICO',
        sizes=[(img.size[0], img.size[1]) for img in images]
    )
    
    print(f"아이콘 파일 생성 완료: {ico_path}")
    return ico_path

if __name__ == '__main__':
    try:
        create_check_icon()
    except Exception as e:
        print(f"아이콘 생성 중 오류 발생: {e}")
        import traceback
        traceback.print_exc()



