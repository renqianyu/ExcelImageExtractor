ExcelImageExtractor
一个简单易用的桌面工具，用于从 Excel 文件（.xlsx）中批量提取插入的图片（位于 xl/media）。
适合不想安装 Python 的普通用户，只需双击即可运行。
可以直接下载ExcelImageExtractor\dist\extract_gui.exe文件
ExcelImageExtractor
A simple and user-friendly desktop tool for batch-extracting images embedded in Excel (.xlsx) files (located in xl/media).
Designed for users who don't want to install Python — just double-click and run.
You can directly download the executable here:
ExcelImageExtractor\dist\extract_gui.exe

✨ 功能特点
✔️ 从 .xlsx 文件中批量提取所有图片
✔️ 自动识别格式：.jpg / .jpeg / .png
✔️ 输出为按顺序编号的图片：image_1.jpg、image_2.png…
✔️ 可视化界面（无需命令行）
✔️ 无依赖、可直接分发给其他人
✔️ 异常报错可视化，便于排查问题
✔️ 提取完成后自动关闭程序

🖼️ 使用方法
打开 ExcelImageExtractor.exe
程序会弹出一个简单的 GUI 窗口。
选择 Excel 文件
点击“浏览”并选择一个 .xlsx 文件。
选择输出文件夹
图片将被导出到你选择的路径中。
点击“开始提取”
程序会从 xl/media/ 提取所有图片
图片按顺序命名
完成后会弹出提示，并自动关闭程序

依赖
仅使用 Python 标准库：
tkinter
zipfile
os
traceback
无需安装第三方库。

🧩 代码说明
主要逻辑：
Excel 本质上是 zip 文件
插入的图片全部存放在 xl/media/
程序直接解析 zip 并导出图片
GUI 使用 tkinter 实现输入与提示功能

✨ Features

✔️ Batch extract all images from .xlsx files
✔️ Automatically detects image formats: .jpg / .jpeg / .png
✔️ Outputs sequentially numbered files: image_1.jpg, image_2.png, etc.
✔️ Simple graphical interface — no command line required
✔️ No dependencies — easy to distribute to non-technical users
✔️ Visual error messages for easier debugging
✔️ Automatically closes after extraction is completed

🖼️ How to Use
1. Open ExcelImageExtractor.exe
A simple GUI window will appear.
2. Select an Excel file
Click Browse and choose any .xlsx file.
3. Select an output folder
All extracted images will be saved to the folder you choose.
4. Click Start Extraction
The program will:
Read the Excel file as a ZIP
Extract all images from xl/media/
Save them with sequential file names
After extraction is completed, a message box will appear and the program will close automatically.
📦 Dependencies
This tool uses only Python standard libraries:
tkinter
zipfile
os
traceback
No third-party packages required.
🧩 Code Overview
Main logic:
Excel (.xlsx) is essentially a ZIP archive
Inserted images are stored under xl/media/
The program reads the ZIP and exports all media files
The GUI is built using tkinter for smooth and simple user interaction
