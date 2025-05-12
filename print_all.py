import fitz  # PyMuPDF：PDFページを画像に変換するためのライブラリ
from PIL import Image, ImageWin  # PIL：画像処理（画像リサイズなど）
import win32print  # Windowsのプリンター操作用
import win32ui     # プリンターに直接描画するためのUI操作
import tkinter as tk  # GUI作成用
from tkinter import filedialog, messagebox, ttk  # GUI部品やメッセージボックス
import os  # ファイル存在チェック、削除など

def select_file():
    file_path.set(filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")]))

def print_pdf():
    path = file_path.get()
    printer = printer_combo.get()
    size_choice = size_combo.get()

    if not os.path.exists(path):
        messagebox.showerror("エラー", "PDFファイルが存在しません。")
        return

    size_map = {
        "6インチ": 6 * 32,  # 432pt
        "7インチ": 7 * 32,  # 504pt
        "8インチ": 8 * 32   # 576pt
    }

    width_pixels = size_map.get(size_choice, 203)

    try:
        doc = fitz.open(path)
        page = doc.load_page(0)
        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
        pix.save("temp_page.png")

        image = Image.open("temp_page.png")
        aspect = image.height / image.width
        resized = image.resize((width_pixels, int(width_pixels * aspect)))

        pdc = win32ui.CreateDC()
        pdc.CreatePrinterDC(printer)
        pdc.StartDoc("PDF Image Print")
        pdc.StartPage()

        dib = ImageWin.Dib(resized)
        dib.draw(pdc.GetHandleOutput(), (0, 0, resized.width, resized.height))

        pdc.EndPage()
        pdc.EndDoc()
        pdc.DeleteDC()

        os.remove("temp_page.png")
        messagebox.showinfo("完了", "印刷が完了しました。")
    except Exception as e:
        messagebox.showerror("印刷エラー", str(e))

# GUIウィンドウ
root = tk.Tk()
root.title("PDF 印刷ツール")

file_path = tk.StringVar()

tk.Label(root, text="PDFファイルを選択:").pack(pady=5)
tk.Entry(root, textvariable=file_path, width=50).pack()
tk.Button(root, text="ファイル選択", command=select_file).pack(pady=5)

tk.Label(root, text="用紙サイズを選択:").pack()
size_combo = ttk.Combobox(root, values=[
    "6インチ",
    "7インチ",
    "8インチ"  # A4幅に収まる最大
])
size_combo.current(0)
size_combo.pack(pady=5)

# 印刷時の幅変換
size_map = {
        "6インチ": 6 * 32,  # 432pt
        "7インチ": 7 * 32,  # 504pt
        "8インチ": 8 * 32   # 576pt
    }
tk.Label(root, text="プリンタを選択:").pack()
printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)
printer_names = [p[2] for p in printers]
printer_combo = ttk.Combobox(root, values=printer_names, width=40)
printer_combo.current(0)
printer_combo.pack(pady=5)

tk.Button(root, text="印刷実行", command=print_pdf).pack(pady=10)

root.mainloop()
