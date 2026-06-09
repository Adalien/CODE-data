import sys, os
sys.stdout.reconfigure(encoding='utf-8')
from pillow_heif import register_heif_opener
from PIL import Image
register_heif_opener()

src_dir = r'C:\Users\admin\OneDrive\桌面\2026應付拍照-截圖'
dst_dir = r'C:\Users\admin\OneDrive\桌面\CODE資料\heic_tmp'
os.makedirs(dst_dir, exist_ok=True)

files = ['IMG_0946','IMG_0947','IMG_0948','IMG_0949','IMG_0950',
         'IMG_0951','IMG_0952','IMG_0953','IMG_0954']

for name in files:
    src = os.path.join(src_dir, name + '.HEIC')
    dst = os.path.join(dst_dir, name + '.jpg')
    if os.path.exists(src):
        img = Image.open(src)
        img.save(dst, 'JPEG', quality=75)
        print(f'✓ {name}.jpg')
    else:
        print(f'✗ 找不到 {name}.HEIC')
