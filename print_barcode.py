import tkinter as tk
from tkinter import messagebox
import datetime
import win32print
import win32con
import threading  # عشان الطباعة ما تعلقش الواجهة
import requests    # المكتبة الجديدة لإرسال الـ HTTP Requests

# -------------------------------------------------------------------
# 1. إعدادات الـ API والطابعة
# -------------------------------------------------------------------
API_URL = "http://192.168.15.8/repaingCard/add-repairing-card-to-print"
API_SECRET = "mySuperSecretPassword123"  # الـ Secret key للـ API
PRINTER_NAME = "Xprinter XP-350B"        # اسم طابعة الباركود


# -------------------------------------------------------------------
# 2. دالة الطباعة (زي ما هي بدون تغيير)
# -------------------------------------------------------------------
def print_raw_tspl_to_xprinter(printer_name, barcode_val, display_val):
    try:
        # **** أبعاد الملصق الواحد (راجعها بدقة) ****
        single_label_width_mm = 45.7
        single_label_height_mm = 12.7

        # **** طول الفراغ بين الملصقات (عمودياً) ****
        vertical_gap_length_mm = 3

        # *** الأبعاد الكلية "للورقة" اللي الطابعة هتشوفها ***
        total_print_width_mm = single_label_width_mm
        total_print_height_mm = (
            single_label_height_mm * 2) + vertical_gap_length_mm

        # *** المسافة الرأسية لبداية الملصق الثاني ***
        offset_for_second_label_mm_vertical = single_label_height_mm + vertical_gap_length_mm
        offset_for_second_label_dots_vertical = int(
            offset_for_second_label_mm_vertical * (203 / 25.4))

        # *** مقدار الإزاحة الأفقية للشمال (0.5 سم = 5 ملم) ***
        horizontal_shift_mm = 5
        horizontal_shift_dots = int(horizontal_shift_mm * (203 / 25.4))

        # *** مقدار الإزاحة الرأسية لتحت (0.2 سم = 2 ملم) ***
        vertical_shift_mm = 2
        vertical_shift_dots = int(vertical_shift_mm * (203 / 25.4))

        # -----------------------------------------------
        # --- الإحداثيات الأساسية قبل الإزاحة لتسهيل التعديل ---
        # -----------------------------------------------
        elfath_x_base = 100
        elfath_y_base = 2
        barcode_x_base = 110
        barcode_y_base = 20
        text_x_base = 130
        text_y_base = 65

        tspl_commands = [
            f"SIZE {total_print_width_mm} mm,{total_print_height_mm} mm\n",
            f"GAP {vertical_gap_length_mm} mm,0 mm\n",
            "CLS\n",
            "DIRECTION 1\n",

            # -----------------------------------------------
            # --- أوامر الطباعة للملصق الأول (العلوي) ---
            # -----------------------------------------------
            f"TEXT {elfath_x_base - horizontal_shift_dots},{elfath_y_base + vertical_shift_dots},\"1\",0,1,1,\"@elfathgroup\"\n",
            f"BARCODE {barcode_x_base - horizontal_shift_dots},{barcode_y_base + vertical_shift_dots},\"128\",40,0,0,3,5,\"{barcode_val}\"\n",
            f"TEXT {text_x_base - horizontal_shift_dots},{text_y_base + vertical_shift_dots},\"2\",0,1,1,\"{display_val}\"\n",

            # -----------------------------------------------
            # --- أوامر الطباعة للملصق الثاني (السفلي) ---
            # -----------------------------------------------
            f"TEXT {elfath_x_base - horizontal_shift_dots},{elfath_y_base + offset_for_second_label_dots_vertical + vertical_shift_dots},\"1\",0,1,1,\"@elfathgroup\"\n",
            f"BARCODE {barcode_x_base - horizontal_shift_dots},{barcode_y_base + offset_for_second_label_dots_vertical + vertical_shift_dots},\"128\",40,0,0,3,5,\"{barcode_val}\"\n",
            f"TEXT {text_x_base - horizontal_shift_dots},{text_y_base + offset_for_second_label_dots_vertical + vertical_shift_dots},\"2\",0,1,1,\"{display_val}\"\n",

            "PRINT 1,1\n"
        ]

        raw_data = "".join(tspl_commands)

        # -----------------------------------------------
        # --- إرسال الأوامر للطابعة
        # -----------------------------------------------
        hPrinter = win32print.OpenPrinter(printer_name)
        try:
            hJob = win32print.StartDocPrinter(
                hPrinter, 1, ("Raw Barcode Print", None, "RAW"))
            try:
                if hJob:
                    win32print.WritePrinter(hPrinter, raw_data.encode('utf-8'))
                    print("تم إرسال أوامر الطباعة بنجاح.")
                else:
                    print("فشل في بدء مهمة الطباعة.")
            finally:
                if hJob:
                    win32print.EndDocPrinter(hPrinter)
        finally:
            win32print.ClosePrinter(hPrinter)

    except Exception as e:
        messagebox.showerror("خطأ في الطباعة",
                             f"حدث خطأ أثناء الطباعة:\n{e}\n\n"
                             "رجاءً تأكد من الآتي:\n"
                             "1. الطابعة متوصلة وشغالة.\n"
                             "2. اسمها في الكود مطابق تماماً لاسمها في الويندوز.\n"
                             "3. تشغيل التطبيق كمسؤول.")

# -------------------------------------------------------------------
# 3. دالة إرسال البيانات للـ API والطباعة
# -------------------------------------------------------------------


def submit_data_and_print():
    owner_name = owner_entry.get()
    device_name = device_entry.get()
    fault_description = fault_entry.get("1.0", tk.END).strip()
    phone_number = phone_number_entry.get()
    attachments = attachments_entry.get()  # رجعنا اسم حقل المرفقات تاني

    if not owner_name or not device_name or not fault_description or not phone_number:
        messagebox.showwarning(
            "بيانات ناقصة", "رجاءً املأ جميع الحقول المطلوبة (اسم صاحب الجهاز، اسم الجهاز، العطل، رقم التليفون).")
        return

    try:
        # توليد باركود جديد للعملية الحالية
        timestamp_ms = int(datetime.datetime.now().timestamp() * 1000)
        barcode_val = str(timestamp_ms)[-4:]
        display_val = barcode_val

        # تجهيز الـ Body بتاع الريكويست (الـ Payload)
        payload = {
            "secret": API_SECRET,
            "name": owner_name,
            "damageName": device_name + " - " + fault_description,
            "barcode": barcode_val,
            "phoneNumber": phone_number,
            "state": False,
            # لو الـ API بتاعك بيدعم حقل للمرفقات، ممكن تبعته هنا
            # "attachments": attachments
        }

        # إرسال الـ POST Request
        response = requests.post(API_URL, json=payload)
        response.raise_for_status()

        # فحص الرد من الـ API
        api_response_data = response.json()
        if response.status_code in [200, 201]:
            messagebox.showinfo(
                "نجاح", f"تم حفظ البيانات بنجاح عن طريق الـ API.\nالرد: {api_response_data}")
        else:
            messagebox.showerror(
                "خطأ في الـ API", f"حدث خطأ من الـ API.\nالحالة: {response.status_code}\nالرد: {api_response_data}")
            return

        # مسح الحقول بعد الحفظ
        owner_entry.delete(0, tk.END)
        device_entry.delete(0, tk.END)
        fault_entry.delete("1.0", tk.END)
        phone_number_entry.delete(0, tk.END)
        attachments_entry.delete(0, tk.END)  # مسح حقل المرفقات

        # تحديث النص بتاع آخر عملية طباعة
        last_print_label.config(
            text=f"آخر باركود مطبوع: {display_val}\nاسم صاحب الجهاز: {owner_name}\nاسم الجهاز: {device_name}")

        # بدء عملية الطباعة في Thread منفصل عشان الواجهة متوقفش
        threading.Thread(target=print_raw_tspl_to_xprinter,
                         args=(PRINTER_NAME, barcode_val, display_val)).start()

    except requests.exceptions.RequestException as e:
        messagebox.showerror("خطأ في الاتصال بالـ API",
                             f"حدث خطأ أثناء الاتصال بالـ API:\n{e}\n\n"
                             "رجاءً تأكد من تشغيل سيرفر الـ API على العنوان {API_URL} وأن العنوان صحيح.")
    except Exception as e:
        messagebox.showerror("خطأ غير متوقع", f"حدث خطأ غير متوقع: {e}")


# -------------------------------------------------------------------
# 4. بناء واجهة Tkinter
# -------------------------------------------------------------------
root = tk.Tk()
root.title("تطبيق إدارة صيانة الأجهزة")
root.geometry("500x600")
root.resizable(False, False)

# تحسينات جمالية
root.configure(bg="#e0f2f7")
root.option_add('*Font', 'Arial 12')

# إطار للادخال
input_frame = tk.LabelFrame(
    root, text="إدخال بيانات الجهاز", padx=20, pady=20, bg="#e0f2f7", fg="#004d40")
input_frame.pack(padx=20, pady=20, fill="both", expand=True)

# اسم صاحب الجهاز
tk.Label(input_frame, text="اسم صاحب الجهاز:", bg="#e0f2f7",
         fg="#004d40").grid(row=0, column=0, sticky="w", pady=5)
owner_entry = tk.Entry(input_frame, width=40)
owner_entry.grid(row=0, column=1, pady=5, padx=10)

# اسم الجهاز
tk.Label(input_frame, text="اسم الجهاز:", bg="#e0f2f7",
         fg="#004d40").grid(row=1, column=0, sticky="w", pady=5)
device_entry = tk.Entry(input_frame, width=40)
device_entry.grid(row=1, column=1, pady=5, padx=10)

# العطل (نص متعدد الأسطر)
tk.Label(input_frame, text="العطل:", bg="#e0f2f7", fg="#004d40").grid(
    row=2, column=0, sticky="nw", pady=5)
fault_entry = tk.Text(input_frame, width=40, height=5)
fault_entry.grid(row=2, column=1, pady=5, padx=10)

# حقل رقم التليفون
tk.Label(input_frame, text="رقم التليفون:", bg="#e0f2f7",
         fg="#004d40").grid(row=3, column=0, sticky="w", pady=5)
phone_number_entry = tk.Entry(input_frame, width=40)
phone_number_entry.grid(row=3, column=1, pady=5, padx=10)

# **التعديل هنا:** رجعنا اسم الحقل "المرفقات"
tk.Label(input_frame, text="المرفقات:", bg="#e0f2f7",
         fg="#004d40").grid(row=4, column=0, sticky="w", pady=5)
attachments_entry = tk.Entry(input_frame, width=40)
attachments_entry.grid(row=4, column=1, pady=5, padx=10)


# زرار الإضافة والطباعة
submit_button = tk.Button(root, text="إضافة وطباعة باركود", command=submit_data_and_print,
                          bg="#00796b", fg="white", activebackground="#004d40", activeforeground="white",
                          relief="raised", bd=3)
submit_button.pack(pady=10)

# عرض آخر نتيجة مطبوعة
last_print_label = tk.Label(root, text="آخر باركود مطبوع: (لا يوجد)",
                            bg="#e0f2f7", fg="#d32f2f", font=('Arial', 10, 'bold'))
last_print_label.pack(pady=10)


root.mainloop()
