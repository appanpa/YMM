import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image
from pillow_heif import register_heif_opener
import pikepdf

register_heif_opener()

SUPPORTED_IMG = ('.jpg', '.jpeg', '.png', '.heic')


def collect_images(folder):
    return [
        os.path.join(folder, f)
        for f in os.listdir(folder)
        if f.lower().endswith(SUPPORTED_IMG)
    ]


def compress_to_jpg(src_path, dst_path, max_mb):
    img = Image.open(src_path).convert('RGB')
    max_bytes = max_mb * 1024 * 1024
    quality = 95
    while quality >= 10:
        img.save(dst_path, 'JPEG', quality=quality, optimize=True)
        if os.path.getsize(dst_path) <= max_bytes:
            break
        quality -= 5


# ── 圖片壓縮頁面 ──────────────────────────────────────────
class ImageTab(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent, bg='#ecf0f1')
        self._build()

    def _build(self):
        f = tk.Frame(self, bg='#ecf0f1', padx=20, pady=16)
        f.pack(fill='both', expand=True)

        tk.Label(f, text='來源資料夾', bg='#ecf0f1').grid(row=0, column=0, sticky='w', pady=6)
        self.src = tk.StringVar()
        tk.Entry(f, textvariable=self.src, width=32).grid(row=0, column=1, padx=6)
        tk.Button(f, text='選擇', command=self._pick_src).grid(row=0, column=2)

        tk.Label(f, text='輸出資料夾', bg='#ecf0f1').grid(row=1, column=0, sticky='w', pady=6)
        self.dst = tk.StringVar()
        tk.Entry(f, textvariable=self.dst, width=32).grid(row=1, column=1, padx=6)
        tk.Button(f, text='選擇', command=self._pick_dst).grid(row=1, column=2)

        tk.Label(f, text='目標大小（MB）', bg='#ecf0f1').grid(row=2, column=0, sticky='w', pady=6)
        self.mb = tk.DoubleVar(value=2.0)
        tk.Spinbox(f, from_=0.1, to=50, increment=0.1,
                   textvariable=self.mb, width=8, format='%.1f').grid(row=2, column=1, sticky='w', padx=6)

        tk.Button(self, text='開始轉換', font=('', 12, 'bold'),
                  bg='#e67e22', fg='white', relief='flat', pady=8,
                  command=self._run).pack(fill='x', padx=20, pady=8)

        self.progress = ttk.Progressbar(self, mode='determinate')
        self.progress.pack(fill='x', padx=20, pady=4)

        self.status = tk.StringVar()
        tk.Label(self, textvariable=self.status, bg='#ecf0f1',
                 wraplength=480, justify='left').pack(padx=20, pady=4)

    def _pick_src(self):
        folder = filedialog.askdirectory(title='選擇來源資料夾')
        if folder:
            self.src.set(folder)
            if not self.dst.get():
                self.dst.set(os.path.join(os.path.dirname(folder), 'converted_photos'))

    def _pick_dst(self):
        folder = filedialog.askdirectory(title='選擇輸出資料夾')
        if folder:
            self.dst.set(folder)

    def _run(self):
        src, dst = self.src.get().strip(), self.dst.get().strip()
        if not src or not os.path.isdir(src):
            messagebox.showerror('錯誤', '請選擇有效的來源資料夾')
            return
        if not dst:
            messagebox.showerror('錯誤', '請選擇輸出資料夾')
            return

        images = collect_images(src)
        if not images:
            messagebox.showinfo('提示', '找不到支援的圖片（JPG / PNG / HEIC）')
            return

        os.makedirs(dst, exist_ok=True)
        self.progress['maximum'] = len(images)
        self.progress['value'] = 0
        errors = []

        for i, path in enumerate(images):
            out_name = os.path.splitext(os.path.basename(path))[0] + '.jpg'
            self.status.set(f'處理中：{os.path.basename(path)}（{i+1}/{len(images)}）')
            self.update()
            try:
                compress_to_jpg(path, os.path.join(dst, out_name), self.mb.get())
            except Exception as e:
                errors.append(f'{os.path.basename(path)}: {e}')
            self.progress['value'] = i + 1
            self.update()

        if errors:
            self.status.set(f'完成（{len(images)-len(errors)} 成功，{len(errors)} 失敗）\n輸出：{dst}')
            messagebox.showwarning('部分失敗', '\n'.join(errors))
        else:
            self.status.set(f'全部完成！共 {len(images)} 張\n輸出：{dst}')
            messagebox.showinfo('完成', f'共 {len(images)} 張轉換完成！\n輸出：{dst}')


# ── PDF 加密/解密頁面 ──────────────────────────────────────
class PdfEncryptTab(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent, bg='#ecf0f1')
        self._build()

    def _build(self):
        f = tk.Frame(self, bg='#ecf0f1', padx=20, pady=16)
        f.pack(fill='both', expand=True)

        # PDF 檔案選擇
        tk.Label(f, text='PDF 檔案', bg='#ecf0f1').grid(row=0, column=0, sticky='w', pady=6)
        self.pdf_path = tk.StringVar()
        tk.Entry(f, textvariable=self.pdf_path, width=36).grid(row=0, column=1, padx=6)
        tk.Button(f, text='選擇', command=self._pick_pdf).grid(row=0, column=2)

        # 操作模式
        tk.Label(f, text='操作', bg='#ecf0f1').grid(row=1, column=0, sticky='w', pady=6)
        self.mode = tk.StringVar(value='encrypt')
        mode_frame = tk.Frame(f, bg='#ecf0f1')
        mode_frame.grid(row=1, column=1, sticky='w', padx=6)
        tk.Radiobutton(mode_frame, text='加密（設定密碼）', variable=self.mode,
                       value='encrypt', bg='#ecf0f1').pack(side='left')
        tk.Radiobutton(mode_frame, text='解密（移除密碼）', variable=self.mode,
                       value='decrypt', bg='#ecf0f1').pack(side='left', padx=12)

        # 密碼
        tk.Label(f, text='密碼', bg='#ecf0f1').grid(row=2, column=0, sticky='w', pady=6)
        self.password = tk.StringVar()
        tk.Entry(f, textvariable=self.password, show='*', width=36).grid(row=2, column=1, padx=6)

        # 輸出路徑
        tk.Label(f, text='輸出檔案', bg='#ecf0f1').grid(row=3, column=0, sticky='w', pady=6)
        self.out_path = tk.StringVar()
        tk.Entry(f, textvariable=self.out_path, width=36).grid(row=3, column=1, padx=6)
        tk.Button(f, text='選擇', command=self._pick_out).grid(row=3, column=2)

        tk.Button(self, text='執行', font=('', 12, 'bold'),
                  bg='#e67e22', fg='white', relief='flat', pady=8,
                  command=self._run).pack(fill='x', padx=20, pady=8)

        self.status = tk.StringVar()
        tk.Label(self, textvariable=self.status, bg='#ecf0f1',
                 wraplength=480, justify='left').pack(padx=20, pady=4)

    def _pick_pdf(self):
        path = filedialog.askopenfilename(title='選擇 PDF 檔案',
                                          filetypes=[('PDF 檔案', '*.pdf')])
        if path:
            self.pdf_path.set(path)
            base = os.path.splitext(path)[0]
            suffix = '_encrypted.pdf' if self.mode.get() == 'encrypt' else '_decrypted.pdf'
            self.out_path.set(base + suffix)

    def _pick_out(self):
        path = filedialog.asksaveasfilename(title='儲存為',
                                             defaultextension='.pdf',
                                             filetypes=[('PDF 檔案', '*.pdf')])
        if path:
            self.out_path.set(path)

    def _run(self):
        src = self.pdf_path.get().strip()
        dst = self.out_path.get().strip()
        pwd = self.password.get()

        if not src or not os.path.isfile(src):
            messagebox.showerror('錯誤', '請選擇有效的 PDF 檔案')
            return
        if not dst:
            messagebox.showerror('錯誤', '請選擇輸出路徑')
            return
        if not pwd:
            messagebox.showerror('錯誤', '請輸入密碼')
            return

        try:
            if self.mode.get() == 'encrypt':
                with pikepdf.open(src) as pdf:
                    pdf.save(dst, encryption=pikepdf.Encryption(
                        owner=pwd, user=pwd, R=6  # R=6 為 256-bit AES
                    ))
                self.status.set(f'加密完成！\n輸出：{dst}')
                messagebox.showinfo('完成', f'加密完成！\n輸出：{dst}')
            else:
                with pikepdf.open(src, password=pwd) as pdf:
                    pdf.save(dst)
                self.status.set(f'解密完成！\n輸出：{dst}')
                messagebox.showinfo('完成', f'解密完成！\n輸出：{dst}')
        except pikepdf.PasswordError:
            messagebox.showerror('錯誤', '密碼錯誤，無法開啟此 PDF')
        except Exception as e:
            messagebox.showerror('錯誤', str(e))


# ── 主視窗 ────────────────────────────────────────────────
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('偵查支援工具箱（本機版）')
        self.geometry('640x440')
        self.resizable(False, False)
        self.configure(bg='#ecf0f1')
        self._build()

    def _build(self):
        tk.Label(self, text='偵查支援工具箱（本機版）', font=('', 15, 'bold'),
                 bg='#2c3e50', fg='white').pack(fill='x', ipady=10)

        notebook = ttk.Notebook(self)
        notebook.pack(fill='both', expand=True, padx=0, pady=0)

        img_tab = ImageTab(notebook)
        notebook.add(img_tab, text='  圖片壓縮  ')

        pdf_tab = PdfEncryptTab(notebook)
        notebook.add(pdf_tab, text='  PDF 加密/解密  ')

        tk.Label(self, text='工具箱 v1.0',
                 bg='#2c3e50', fg='#7f8c8d', font=('', 10)).pack(
            fill='x', side='bottom', ipady=4)


if __name__ == '__main__':
    App().mainloop()
