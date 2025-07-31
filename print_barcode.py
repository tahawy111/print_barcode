import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import messagebox, END, VERTICAL, X, W, EW, NS, N, S, LEFT, RIGHT
import datetime
import win32print
import win32api
import threading
import requests
import webbrowser

# -------------------------------------------------------------------
# 1. إعدادات الـ API والطابعة
# -------------------------------------------------------------------
API_URL = "http://192.168.15.8/repaingCard/add-repairing-card-to-print"
PRINT_RECEIPT_BASE_URL = "http://192.168.15.8/repaingCard/print"
API_SECRET = "mySuperSecretPassword123"
PRINTER_NAME = "Xprinter XP-350B"


# -------------------------------------------------------------------
# 2. دالة الطباعة (بدون تغيير جوهري)
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

            # --- أوامر الطباعة للملصق الأول (العلوي) ---
            f"TEXT {elfath_x_base - horizontal_shift_dots},{elfath_y_base + vertical_shift_dots},\"1\",0,1,1,\"@elfathgroup\"\n",
            f"BARCODE {barcode_x_base - horizontal_shift_dots},{barcode_y_base + vertical_shift_dots},\"128\",40,0,0,3,5,\"{barcode_val}\"\n",
            f"TEXT {text_x_base - horizontal_shift_dots},{text_y_base + vertical_shift_dots},\"2\",0,1,1,\"{display_val}\"\n",

            # --- أوامر الطباعة للملصق الثاني (السفلي) ---
            f"TEXT {elfath_x_base - horizontal_shift_dots},{elfath_y_base + offset_for_second_label_dots_vertical + vertical_shift_dots},\"1\",0,1,1,\"@elfathgroup\"\n",
            f"BARCODE {barcode_x_base - horizontal_shift_dots},{barcode_y_base + offset_for_second_label_dots_vertical + vertical_shift_dots},\"128\",40,0,0,3,5,\"{barcode_val}\"\n",
            f"TEXT {text_x_base - horizontal_shift_dots},{text_y_base + offset_for_second_label_dots_vertical + vertical_shift_dots},\"2\",0,1,1,\"{display_val}\"\n",

            "PRINT 1,1\n"
        ]

        raw_data = "".join(tspl_commands)

        # --- إرسال الأوامر للطابعة ---
        hPrinter = win32print.OpenPrinter(printer_name)
        try:
            hJob = win32print.StartDocPrinter(
                hPrinter, 1, ("Raw Barcode Print", None, "RAW"))
            try:
                if hJob:
                    win32print.WritePrinter(hPrinter, raw_data.encode('utf-8'))
                else:
                    messagebox.showerror(
                        "خطأ طباعة", "فشل في بدء مهمة الطباعة.")
            finally:
                if hJob:
                    win32print.EndDocPrinter(hPrinter)
        finally:
            win32print.ClosePrinter(hPrinter)

    except Exception as e:
        messagebox.showerror("خطأ في الطباعة",
                             f"حدث خطأ أثناء الطباعة:\n{e}\n\n"
                             "تأكد من:\n"
                             "1. الطابعة متصلة وتعمل.\n"
                             "2. اسم الطابعة في الكود مطابق لاسمها في الويندوز.\n"
                             "3. تشغيل التطبيق كمسؤول (Administrator).")

# -------------------------------------------------------------------
# 3. دوال الأزرار الجديدة
# -------------------------------------------------------------------


def reprint_barcode_action():
    """تعيد طباعة الباركود الأخير الذي تم إنشاؤه."""
    if root.last_barcode_val and root.last_display_val:
        status_var.set("جاري إعادة طباعة الباركود...")
        print_thread = threading.Thread(target=print_raw_tspl_to_xprinter, args=(
            PRINTER_NAME, root.last_barcode_val, root.last_display_val))
        print_thread.start()
    else:
        messagebox.showwarning(
            "لا توجد بيانات", "لا توجد بيانات باركود سابقة لإعادة طباعتها.")


def print_receipt_action():
    """تفتح صفحة إيصال الصيانة في المتصفح الافتراضي."""
    if root.last_repair_card_id:
        receipt_url = f"{PRINT_RECEIPT_BASE_URL}/{root.last_repair_card_id}"
        status_var.set(f"فتح الإيصال في المتصفح: {receipt_url}")
        webbrowser.open_new_tab(receipt_url)
    else:
        messagebox.showwarning(
            "لا توجد بيانات", "لا توجد بيانات كرت صيانة سابقة لطباعة إيصالها.")

# -------------------------------------------------------------------
# 4. دالة إرسال البيانات للـ API والطباعة (مع التعديلات)
# -------------------------------------------------------------------


def submit_data_and_print():
    status_var.set("جاري الإرسال والطباعة...")
    root.update_idletasks()
    submit_button.config(state="disabled")
    reprint_barcode_button.config(state="disabled")
    print_receipt_button.config(state="disabled")

    owner_name = owner_entry.get()
    device_name = device_entry.get()
    fault_description = fault_entry.get().strip()  # استخدام .get() لحقل input عادي
    phone_number = phone_number_entry.get()
    attachments = attachments_entry.get()
    cost_value = cost_entry.get()

    if not all([owner_name, device_name, fault_description, phone_number]):
        messagebox.showwarning(
            "بيانات ناقصة", "يرجى ملء جميع الحقول المطلوبة (الاسم، الجهاز، العطل، رقم الهاتف).")
        status_var.set("في انتظار الإدخال...")
        submit_button.config(state="normal")
        return

    try:
        timestamp_ms = int(datetime.datetime.now().timestamp() * 1000)
        barcode_val = str(timestamp_ms)[-4:]
        display_val = barcode_val

        payload = {
            "secret": API_SECRET,
            "name": owner_name,
            "damageName": f"{device_name} - {fault_description}",
            "barcode": barcode_val,
            "phoneNumber": phone_number,
            "state": False,
        }

        # إضافة التكلفة إذا كانت موجودة (يتم إرسالها مع نفس الريكويست)
        if cost_value:
            try:
                payload["cost"] = float(cost_value)
            except ValueError:
                messagebox.showwarning(
                    "خطأ في التكلفة", "يرجى إدخال قيمة رقمية صحيحة للتكلفة أو تركها فارغة.")
                status_var.set("خطأ في إدخال التكلفة.")
                submit_button.config(state="normal")
                return
        else:
            # إرسال 0 أو يمكن إرسال None حسب تصميم الـ API بتاعك
            payload["cost"] = 0

        response = requests.post(API_URL, json=payload, timeout=10)
        response.raise_for_status()

        api_response_data = response.json()
        if response.status_code in [200, 201]:
            messagebox.showinfo(
                "نجاح", f"تم حفظ البيانات بنجاح.\nالرد: {api_response_data.get('message', 'لا يوجد رسالة')}")

            # تخزين البيانات لزرار إعادة الطباعة والإيصال
            root.last_barcode_val = barcode_val
            root.last_display_val = display_val
            root.last_repair_card_id = api_response_data.get('cardId')

            # تفعيل أزرار إعادة الطباعة والإيصال
            reprint_barcode_button.config(state="normal")
            print_receipt_button.config(state="normal")

        else:
            messagebox.showerror(
                "خطأ API", f"حدث خطأ من الـ API.\nالحالة: {response.status_code}\nالرد: {api_response_data.get('message', 'لا يوجد رسالة')}")
            status_var.set("فشل إرسال البيانات.")
            submit_button.config(state="normal")
            return

        # مسح الحقول بعد النجاح
        for entry in [owner_entry, device_entry, phone_number_entry, attachments_entry, cost_entry, fault_entry]:
            entry.delete(0, END)

        last_print_label.config(
            text=f"آخر عملية: {display_val} | {owner_name} | {device_name}")

        # بدء الطباعة في Thread منفصل (للبراكود)
        print_thread = threading.Thread(target=print_raw_tspl_to_xprinter, args=(
            PRINTER_NAME, barcode_val, display_val))
        print_thread.start()

        status_var.set("تمت العملية بنجاح. في انتظار إدخال جديد...")

    except requests.exceptions.RequestException as e:
        messagebox.showerror(
            "خطأ اتصال", f"فشل الاتصال بالـ API:\n{e}\n\nتأكد من تشغيل السيرفر على العنوان الصحيح.")
        status_var.set("خطأ في الاتصال بالـ API.")
    except Exception as e:
        messagebox.showerror("خطأ غير متوقع", f"حدث خطأ غير متوقع: {e}")
        status_var.set("حدث خطأ غير متوقع.")
    finally:
        # التأكد من إعادة تفعيل الزر في كل الحالات
        submit_button.config(state="normal")
        owner_entry.focus_set()


# -------------------------------------------------------------------
# 5. بناء واجهة ttkbootstrap الرسومية
# -------------------------------------------------------------------

# --- إعداد النافذة الرئيسية ---
root = ttk.Window(themename="litera")
root.title("برنامج إدارة الصيانة")
root.geometry("600x750")  # زيادة ارتفاع النافذة لاستيعاب الأزرار الجديدة
root.resizable(False, False)

# لتعيين الأيقونة، تأكد من وجود ملف "icon.ico" في نفس المجلد
try:
    root.iconbitmap("icon.ico")
except:
    print("لم يتم العثور على ملف الأيقونة icon.ico")

# --- دالة لإنشاء قائمة النسخ واللصق ---


def create_context_menu(widget):
    menu = ttk.Menu(widget, tearoff=0)
    menu.add_command(
        label="قص", command=lambda: widget.event_generate("<<Cut>>"))
    menu.add_command(
        label="نسخ", command=lambda: widget.event_generate("<<Copy>>"))
    menu.add_command(
        label="لصق", command=lambda: widget.event_generate("<<Paste>>"))
    menu.add_separator()
    menu.add_command(label="تحديد الكل",
                     command=lambda: widget.event_generate("<<SelectAll>>"))

    def show_menu(event):
        widget.focus_set()
        menu.post(event.x_root, event.y_root)

    widget.bind("<Button-3>", show_menu)


# --- إطار إدخال البيانات ---
input_frame = ttk.LabelFrame(
    root, text="بيانات الجهاز الجديد", padding=(20, 15))
input_frame.pack(padx=20, pady=20, fill=X, anchor=N)
input_frame.columnconfigure(1, weight=1)

# --- حقول الإدخال ---
labels_texts = ["اسم صاحب الجهاز:", "اسم الجهاز:",
                "وصف العطل:", "رقم التليفون:", "المرفقات (اختياري):", "التكلفة (اختياري):"]
entries = {}

for i, text in enumerate(labels_texts):
    label = ttk.Label(input_frame, text=text)
    label.grid(row=i, column=0, sticky=W, padx=5, pady=10)

    if text == "وصف العطل:":
        entry = ttk.Entry(input_frame, width=40, font=(
            "Arial", 11))  # تم التغيير لـ ttk.Entry
        entry.grid(row=i, column=1, sticky=EW, padx=5, pady=10)
        # تم حذف كود Scrollbar
        fault_entry = entry
    else:
        entry = ttk.Entry(input_frame, width=40, font=("Arial", 11))
        entry.grid(row=i, column=1, sticky=EW, padx=5, pady=10)
        if text == "التكلفة (اختياري):":
            cost_entry = entry

    create_context_menu(entry)  # تفعيل قائمة السياق لكل حقل
    entries[text] = entry

# إعادة تعيين أسماء المتغيرات الأصلية لتعمل الدوال كما هي
owner_entry = entries[labels_texts[0]]
device_entry = entries[labels_texts[1]]
phone_number_entry = entries[labels_texts[3]]
attachments_entry = entries[labels_texts[4]]


# --- زر الإضافة والطباعة ---
submit_button = ttk.Button(
    root,
    text="إضافة وطباعة باركود",
    command=submit_data_and_print,
    bootstyle=(PRIMARY, OUTLINE),  # شكل مميز للزر
    padding=(20, 10)
)
submit_button.pack(pady=(0, 10))  # مسافة أقل بين الأزرار


# --- أزرار الطباعة الجديدة ---
reprint_barcode_button = ttk.Button(
    root,
    text="إعادة طباعة الباركود",
    command=reprint_barcode_action,
    bootstyle=(INFO, OUTLINE),  # لون مختلف عشان يبان
    padding=(15, 8),
    state="disabled"  # معطل مبدئياً
)
reprint_barcode_button.pack(pady=(0, 10))

print_receipt_button = ttk.Button(
    root,
    text="طباعة إيصال الصيانة",
    command=print_receipt_action,
    bootstyle=(SUCCESS, OUTLINE),  # لون مختلف عشان يبان
    padding=(15, 8),
    state="disabled"  # معطل مبدئياً
)
print_receipt_button.pack(pady=(0, 20))


# --- إطار عرض الحالة ---
status_frame = ttk.LabelFrame(root, text="الحالة", padding=(10, 10))
status_frame.pack(padx=20, pady=10, fill=X, anchor=S)

last_print_label = ttk.Label(status_frame, text="آخر عملية: (لا يوجد)", font=(
    "Arial", 10), bootstyle=SECONDARY)
last_print_label.pack(side=LEFT, padx=5)

status_var = ttk.StringVar(value="جاهز لاستقبال البيانات...")
status_label = ttk.Label(status_frame, textvariable=status_var, font=(
    "Arial", 9, "italic"), bootstyle=INFO)
status_label.pack(side=RIGHT, padx=5)

# --- متغيرات لتخزين آخر بيانات تم استخدامها للطباعة ---
root.last_barcode_val = None
root.last_display_val = None
root.last_repair_card_id = None  # لتخزين الـ ID بتاع الكارت بعد الحفظ

# --- بدء تشغيل الواجهة ---
owner_entry.focus_set()  # التركيز على أول حقل عند بدء التشغيل
root.mainloop()
