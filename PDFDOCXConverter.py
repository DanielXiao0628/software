import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from pdf2docx import Converter
from docx2pdf import convert
import webbrowser

def pdf_to_docx():
    file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if file_path:
        docx_output_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("DOCX files", "*.docx")])
        if docx_output_path:
            cv = Converter(file_path)
            cv.convert(docx_output_path)
            cv.close()
            messagebox.showinfo("Conversion Successful", "PDF to DOCX conversion completed!")
            webbrowser.open("https://majorcharacter.com/bC3.VS0/PD3gp/vBbEmvVEJLZKDb0_0iOADrUN5HOiDaE/xsLsTyQC4vNNTAkT4EMNT_IE")
            webbrowser.open("https://fostmar.online/")
def docx_to_pdf():
    file_path = filedialog.askopenfilename(filetypes=[("DOCX files", "*.docx")])
    if file_path:
        pdf_output_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if pdf_output_path:
            convert(file_path, pdf_output_path)
            messagebox.showinfo("Conversion Successful", "DOCX to PDF conversion completed!")
            webbrowser.open("https://majorcharacter.com/bC3.VS0/PD3gp/vBbEmvVEJLZKDb0_0iOADrUN5HOiDaE/xsLsTyQC4vNNTAkT4EMNT_IE")
            webbrowser.open("https://fostmar.online/")

def center_window(window, width, height):
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    x = (screen_width - width) // 2
    y = (screen_height - height) // 2
    window.geometry(f"{width}x{height}+{x}+{y}")

# 创建主窗口
root = tk.Tk()
root.title("PDF/DOCX Converter")

# 设置窗口大小
root.geometry("400x200")
center_window(root, 400, 100)

# 创建按钮，增大按钮的宽度
pdf_to_docx_button = tk.Button(root, text="PDF to DOCX", command=pdf_to_docx, width=50)
docx_to_pdf_button = tk.Button(root, text="DOCX to PDF", command=docx_to_pdf, width=50)

# 安排按钮在主窗口
pdf_to_docx_button.pack(pady=10)
docx_to_pdf_button.pack(pady=10)

# 启动主循环
root.mainloop()
