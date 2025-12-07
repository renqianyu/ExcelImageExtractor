import os
import zipfile
import tkinter as tk
from tkinter import filedialog, messagebox
import traceback

# ---------- 图片提取功能 ----------
def extract_images(excel_path, output_dir):
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    with zipfile.ZipFile(excel_path, 'r') as z:
        # 找所有 media 文件
        media_files = [f for f in z.namelist() if f.lower().startswith("xl/media/") 
                       and f.lower().endswith(('.jpeg', '.jpg', '.png'))]
        
        if not media_files:
            messagebox.showwarning("提示", "⚠️ 未找到 xl/media 下的图片文件")
            return

        for idx, media_file in enumerate(media_files, start=1):
            data = z.read(media_file)
            ext = os.path.splitext(media_file)[1]
            save_name = f"image_{idx}{ext}"
            save_path = os.path.join(output_dir, save_name)
            with open(save_path, 'wb') as f:
                f.write(data)

    messagebox.showinfo("完成", f"✅ 全部图片已导出到:\n{output_dir}")

# ---------- GUI 部分 ----------
class ExcelExtractorApp:
    def __init__(self, master):
        self.master = master
        master.title("Excel 图片提取器")
        master.geometry("480x250")
        master.resizable(False, False)

        # Excel 文件选择
        tk.Label(master, text="选择 Excel 文件:").pack(anchor="w", padx=20, pady=(10,0))
        self.excel_entry = tk.Entry(master, width=60)
        self.excel_entry.pack(padx=20)
        tk.Button(master, text="浏览", command=self.browse_excel, width=12).pack(pady=(5,10))

        # 输出文件夹选择
        tk.Label(master, text="选择输出文件夹:").pack(anchor="w", padx=20, pady=(5,0))
        self.output_entry = tk.Entry(master, width=60)
        self.output_entry.pack(padx=20)
        tk.Button(master, text="浏览", command=self.browse_output, width=12).pack(pady=(5,10))

        # 开始提取按钮
        tk.Button(master, text="开始提取", command=self.start_extraction,
                  bg="#4CAF50", fg="white", width=25, height=2).pack(pady=15)

    def browse_excel(self):
        file = filedialog.askopenfilename(filetypes=[("Excel 文件", "*.xlsx")])
        if file:
            self.excel_entry.delete(0, tk.END)
            self.excel_entry.insert(0, file)

    def browse_output(self):
        folder = filedialog.askdirectory()
        if folder:
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, folder)

    def start_extraction(self):
        excel_path = self.excel_entry.get()
        output_dir = self.output_entry.get()

        if not excel_path or not os.path.isfile(excel_path):
            messagebox.showerror("错误", "请选择有效的 Excel 文件")
            return
        if not output_dir:
            messagebox.showerror("错误", "请选择输出文件夹")
            return

        try:
            extract_images(excel_path, output_dir)
            # 提取完成后自动关闭窗口
            self.master.destroy()
        except Exception:
            err_msg = traceback.format_exc()
            messagebox.showerror("提取错误", f"提取过程中出现错误:\n{err_msg}")

# ---------- 主程序 ----------
if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelExtractorApp(root)
    root.mainloop()
