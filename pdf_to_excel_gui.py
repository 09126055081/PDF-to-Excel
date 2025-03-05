import tkinter as tk
from tkinter import filedialog, messagebox
import pdfplumber
import pandas as pd
import os
import subprocess

excel_file_path = ""  # مسیر فایل خروجی Excel رو ذخیره می‌کنیم

def pdf_to_excel(pdf_file):
    global excel_file_path  
    try:
        with pdfplumber.open(pdf_file) as pdf:
            data = []  # برای ذخیره همه اطلاعات PDF

            for page_num, page in enumerate(pdf.pages, start=1):
                text = page.extract_text()  # استخراج متن هر صفحه
                if text:
                    data.append(["--- صفحه " + str(page_num) + " ---"])  # جداسازی صفحات در اکسل
                    for line in text.split("\n"):
                        data.append([line])  # هر خط رو به لیست اضافه کن

                tables = page.extract_tables()
                for table in tables:
                    for row in table:
                        data.append(row)  # اضافه کردن ردیف‌های جدول

            df = pd.DataFrame(data)

            # ذخیره در یک فایل اکسل در همان مسیر PDF
            excel_file_path = os.path.splitext(pdf_file)[0] + "_converted.xlsx"
            df.to_excel(excel_file_path, index=False, header=False, engine='openpyxl')

            messagebox.showinfo("موفقیت", f"تبدیل انجام شد!\nفایل ذخیره شد در:\n{excel_file_path}")
            btn_open_excel.config(state=tk.NORMAL)  # فعال کردن دکمه باز کردن فایل

    except Exception as e:
        messagebox.showerror("خطا", f"مشکلی پیش آمد:\n{e}")

def select_pdf():
    file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    if file_path:
        pdf_to_excel(file_path)

def open_excel():
    if os.path.exists(excel_file_path):
        subprocess.run(["start", excel_file_path], shell=True)  # باز کردن فایل در ویندوز
    else:
        messagebox.showerror("خطا", "فایل Excel پیدا نشد!")

# رابط گرافیکی با Tkinter
root = tk.Tk()
root.title("PDF به Excel")
root.geometry("300x200")

label = tk.Label(root, text="انتخاب فایل PDF برای تبدیل", font=("Arial", 12))
label.pack(pady=10)

btn_select = tk.Button(root, text="انتخاب فایل", command=select_pdf, font=("Arial", 12))
btn_select.pack(pady=5)

btn_open_excel = tk.Button(root, text="باز کردن فایل EXCEL", command=open_excel, font=("Arial", 12), state=tk.DISABLED)
btn_open_excel.pack(pady=5)

root.mainloop()
