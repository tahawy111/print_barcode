import win32print
import win32con
import datetime


def print_raw_tspl_to_xprinter(printer_name="Xprinter XP-350B"):
    try:
        # 1. تجهيز بيانات الباركود والرقم اللي هيظهر
        timestamp_ms = int(datetime.datetime.now().timestamp() * 1000)

        # قيمة الباركود هتكون آخر 4 أرقام من الـ timestamp
        barcode_value_for_encoding = str(timestamp_ms)[-4:]

        # الرقم اللي هيظهر كنص تحت الباركود (برضه آخر 4 أرقام)
        display_text_value = str(timestamp_ms)[-4:]

        # **** أبعاد الملصق الواحد اللي انت قستها (راجعها بدقة) ****
        single_label_width_mm = 45.7    # العرض الفعلي للملصق الواحد بالمليمتر
        single_label_height_mm = 12.7   # الارتفاع الفعلي للملصق الواحد بالمليمتر

        # **** طول الفراغ بين الملصقات (عمودياً) بالمليمتر ****
        # المسافة بين الملصق الأول (العلوي) والثاني (السفلي)
        vertical_gap_length_mm = 3

        # *** الأبعاد الكلية "للورقة" اللي الطابعة هتشوفها عشان تطبع عليها ***
        total_print_width_mm = single_label_width_mm
        total_print_height_mm = (
            single_label_height_mm * 2) + vertical_gap_length_mm

        # *** المسافة الرأسية لبداية الملصق الثاني (بالمليمتر) ***
        offset_for_second_label_mm_vertical = single_label_height_mm + vertical_gap_length_mm
        offset_for_second_label_dots_vertical = int(
            offset_for_second_label_mm_vertical * (203 / 25.4))

        # *** مقدار الإزاحة الأفقية للشمال (0.5 سم = 5 ملم) ***
        # هنطرح 5 ملم (وما يعادلها بالدوتس) من كل إحداثيات X
        horizontal_shift_mm = 5
        horizontal_shift_dots = int(horizontal_shift_mm * (203 / 25.4))

        # *** مقدار الإزاحة الرأسية لتحت (0.2 سم = 2 ملم) ***
        # **التعديل هنا:** هنضيف 2 ملم (وما يعادلها بالدوتس) على كل إحداثيات Y
        vertical_shift_mm = 2  # 0.2 سم = 2 ملم
        vertical_shift_dots = int(
            vertical_shift_mm * (203 / 25.4))  # حوالي 16 دوتس

        # -----------------------------------------------
        # --- الإحداثيات الأساسية قبل الإزاحة لتسهيل التعديل ---
        # -----------------------------------------------
        # دي إحداثيات العناصر الأساسية للملصق الواحد (اللي كانت قبل أي إزاحة)
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

            # 1. طباعة "@elfathgroup" مصغر جداً (مع إزاحة للشمال ولتحت)
            f"TEXT {elfath_x_base - horizontal_shift_dots},{elfath_y_base + vertical_shift_dots},\"1\",0,1,1,\"@elfathgroup\"\n",

            # 2. أمر طباعة الباركود Code 128 (مع إزاحة للشمال ولتحت)
            f"BARCODE {barcode_x_base - horizontal_shift_dots},{barcode_y_base + vertical_shift_dots},\"128\",40,0,0,3,5,\"{barcode_value_for_encoding}\"\n",


            # 3. أمر طباعة الرقم الخارجي (مع إزاحة للشمال ولتحت)
            f"TEXT {text_x_base - horizontal_shift_dots},{text_y_base + vertical_shift_dots},\"2\",0,1,1,\"{display_text_value}\"\n",


            # -----------------------------------------------
            # --- أوامر الطباعة للملصق الثاني (السفلي) ---
            # -----------------------------------------------

            # 1. طباعة "@elfathgroup" مصغر جداً في الملصق الثاني (مع إزاحة للشمال ولتحت)
            f"TEXT {elfath_x_base - horizontal_shift_dots},{elfath_y_base + offset_for_second_label_dots_vertical + vertical_shift_dots},\"1\",0,1,1,\"@elfathgroup\"\n",

            # 2. أمر طباعة الباركود Code 128 في الملصق الثاني (مع إزاحة للشمال ولتحت)
            f"BARCODE {barcode_x_base - horizontal_shift_dots},{barcode_y_base + offset_for_second_label_dots_vertical + vertical_shift_dots},\"128\",40,0,0,3,5,\"{barcode_value_for_encoding}\"\n",

            # 3. أمر طباعة الرقم الخارجي في الملصق الثاني (مع إزاحة للشمال ولتحت)
            f"TEXT {text_x_base - horizontal_shift_dots},{text_y_base + offset_for_second_label_dots_vertical + vertical_shift_dots},\"2\",0,1,1,\"{display_text_value}\"\n",


            "PRINT 1,1\n"
        ]

        raw_data = "".join(tspl_commands)

        print(f"جاري إرسال أوامر TSPL للطابعة: {printer_name}...")
        print("أوامر TSPL المرسلة:\n" + raw_data)

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
        print(f"حدث خطأ أثناء الطباعة: {e}")
        print("رجاءً تأكد من الآتي:")
        print("1. الطابعة Xprinter XP-350B متوصلة وشغالة.")
        print("2. اسم الطابعة في الكود مطابق تماماً لاسمها في الويندوز.")
        print("3. تشغيل السكريبت كمسؤول (Run as administrator).")
        print("4. مقاسات الملصق (SIZE و GAP) في أوامر TSPL مظبوطة بدقة مع مقاسات ملصقاتك.")
        print("5. الطابعة فعلاً تدعم أوامر TSPL (غالباً Xprinter بتدعمها).")
        print("6. إحداثيات الباركود والنص (X,Y) وعرض الخطوط (narrow, wide) مناسبة لحجم الملصق.")


if __name__ == "__main__":
    printer_name = "Xprinter XP-350B"
    print_raw_tspl_to_xprinter(printer_name)
