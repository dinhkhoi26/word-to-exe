import os
import subprocess
import shutil
import sys
import time
from pdf2image import convert_from_path
from tkinter import Tk, filedialog

# === Hàm chọn file Word ===
def select_docx_file():
    root = Tk()
    root.withdraw()
    return filedialog.askopenfilename(
        title="🔍 Chọn file Word",
        filetypes=[("Word Documents", "*.docx")]
    )

# === Chuyển Word → PDF bằng LibreOffice ===
def convert_docx_to_pdf(docx_path, output_dir):
    pdf_filename = os.path.splitext(os.path.basename(docx_path))[0] + ".pdf"
    pdf_path = os.path.join(output_dir, pdf_filename)
    libreoffice_path = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
    subprocess.run([
        libreoffice_path,
        "--headless",
        "--convert-to", "pdf",
        "--outdir", output_dir,
        docx_path
    ], check=True)
    return pdf_path

# === Chuyển PDF → PNG ===
def convert_pdf_to_png(pdf_path, output_dir):
    images = convert_from_path(pdf_path, dpi=300)
    image_paths = []
    for i, img in enumerate(images):
        img_path = os.path.join(output_dir, f"page_{i+1}.png")
        img.save(img_path, "PNG")
        image_paths.append(img_path)
    return image_paths

# === Gộp toàn bộ chuyển Word → PNG ===
def convert_docx_to_png(docx_path, output_dir):
    os.makedirs(output_dir, exist_ok=True)
    pdf_path = convert_docx_to_pdf(docx_path, output_dir)
    image_paths = convert_pdf_to_png(pdf_path, output_dir)
    os.remove(pdf_path)
    return image_paths

# === Tạo script Python để in ảnh và tự xóa ===
def create_printer_script(image_path, output_dir):
    script = f'''
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
del "{{exe_path}}"
""")
    os.system(f'start /min cmd /c "{{bat_path}}"')

def main():
    image_path = os.path.join(os.path.dirname(sys.argv[0]), "{os.path.basename(image_path)}")
    if not os.path.exists(image_path):
        ctypes.windll.user32.MessageBoxW(0, "Không tìm thấy ảnh!", "Lỗi", 0)
        return
    show_image(image_path)
    print_image(image_path)
    time.sleep(5)
    delete_self()

if __name__ == "__main__":
    main()
'''.strip()

    script_path = os.path.join(output_dir, "printer.py")
    with open(script_path, "w", encoding="utf-8") as f:
        f.write(script)
    return script_path

# === Đóng gói bằng PyInstaller ===
def build_exe(script_path):
    if shutil.which("pyinstaller") is None:
        print("❌ PyInstaller chưa được cài. Chạy: pip install pyinstaller")
        return
    subprocess.run(["pyinstaller", "--onefile", "--noconsole", script_path], check=True)

# === MAIN ===
def main():
    print("🔍 Chọn file Word...")
    docx_file = select_docx_file()
    if not docx_file:
        print("❌ Không chọn file nào.")
        return

    output_dir = os.path.join(os.path.dirname(docx_file), "output")
    print("🖼️  Chuyển sang ảnh...")
    images = convert_docx_to_png(docx_file, output_dir)

    print("📝 Tạo script in và tự xoá...")
    script_path = create_printer_script(images[0], output_dir)

    print("📦 Đóng gói .exe bằng PyInstaller...")
    os.chdir(output_dir)
    build_exe(script_path)

    print("✅ Hoàn tất! File .exe nằm trong thư mục:")
    print(os.path.join(output_dir, "dist"))

if __name__ == "__main__":
    main()
