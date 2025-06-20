import os
import sys
import time
import ctypes
import win32api
from PIL import Image

def print_image(image_path):
    win32api.ShellExecute(0, "print", image_path, None, ".", 0)

def show_image(image_path):
    img = Image.open(image_path)
    img.show()

def delete_self():
    exe_path = os.path.abspath(sys.argv[0])
    bat_path = exe_path + ".bat"
    with open(bat_path, "w") as f:
        f.write(f"""@echo off
timeout /t 5 >nul
del "%~f0"
del "{exe_path}"
""")
    os.system(f'start /min cmd /c "{bat_path}"')

def main():
    image_path = os.path.join(os.path.dirname(sys.argv[0]), "page_1.png")
    if not os.path.exists(image_path):
        ctypes.windll.user32.MessageBoxW(0, "Không tìm thấy ảnh!", "Lỗi", 0)
        return
    show_image(image_path)
    print_image(image_path)
    time.sleep(5)
    delete_self()

if __name__ == "__main__":
    main()