# وارد کردن کتابخانه‌های مورد نیاز
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QFileDialog, QProgressBar, QLabel, QGridLayout
from PyQt5.QtCore import Qt, QThread, pyqtSignal
import sys
import pandas as pd
import os

# تعریف یک کلاس برای اجرای جستجو در پس‌زمینه
class SearchThread(QThread):
    # تعریف یک سیگنال برای ارسال پیشرفت جستجو به پنجره اصلی
    progress = pyqtSignal(int)

    # تعریف تابع سازنده کلاس
    def __init__(self, excel_file, pdf_folder, save_folder):
        super().__init__()
        # ذخیره مقادیر ورودی در متغیرهای کلاس
        self.excel_file = excel_file
        self.pdf_folder = pdf_folder
        self.save_folder = save_folder
        # ایجاد یک لیست خالی برای ذخیره نام‌های فایل‌هایی که پیدا نشدند
        self.not_found = []

    # تعریف تابع اصلی کلاس که در هنگام شروع رشته اجرا می‌شود
    def run(self):
        # خواندن فایل اکسل و دریافت ستون اول به عنوان یک سری
        data = pd.read_excel(self.excel_file, header=None)
        names = data[0]

        # تعریف یک شمارنده برای محاسبه پیشرفت جستجو
        count = 0
        # حلقه بر روی هر نام در سری
        for name in names:
            # افزایش شمارنده
            count += 1
            # ارسال درصد پیشرفت جستجو به پنجره اصلی با استفاده از سیگنال
            self.progress.emit(int(count / len(names) * 100))
            # تعریف یک پرچم برای بررسی اینکه آیا فایلی با این نام پیدا شده است یا خیر
            found = False
            # حلقه بر روی همه فایل‌های موجود در پوشه PDF و زیرپوشه‌های آن
            for root, dirs, files in os.walk(self.pdf_folder):
                # حلقه بر روی هر فایل
                for file in files:
                    # بررسی اینکه آیا فایل یک فایل PDF است
                    if file.endswith(".pdf"):
                        # بررسی اینکه آیا نام فایل شامل نام مورد جستجو است
                        if name in file:
                            # تعیین مسیر کامل فایل
                            file_path = os.path.join(root, file)
                            # کپی کردن فایل به پوشه ذخیره‌سازی
                            os.system(f"copy {file_path} {self.save_folder}")
                            # تنظیم پرچم به True
                            found = True
                            # خروج از حلقه فایل‌ها
                            break
                # اگر پرچم True باشد
                if found:
                    # خروج از حلقه پوشه‌ها
                    break
            # اگر پرچم False باشد
            if not found:
                # افزودن نام به لیست نام‌های پیدا نشده
                self.not_found.append(name)

        # اگر لیست نام‌های پیدا نشده خالی نباشد
        if self.not_found:
            # تبدیل لیست به یک سری
            not_found_series = pd.Series(self.not_found)
            # ذخیره سری به عنوان یک فایل اکسل با نام Not Found List
            not_found_series.to_excel(os.path.join(self.save_folder, "Not Found List.xlsx"), header=False, index=False)

# تعریف یک کلاس برای پنجره اصلی برنامه
class MainWindow(QWidget):
    # تعریف تابع سازنده کلاس
    def __init__(self):
        super().__init__()
        # تنظیم عنوان پنجره
        self.setWindowTitle("PDF Searcher")
        # تنظیم اندازه پنجره
        self.resize(400, 200)
        # ایجاد چهار دکمه با عناوین مشخص
        self.list_of_file_button = QPushButton("List Of File")
        self.source_pdf_address_button = QPushButton("Source PDF Address")
        self.save_data_button = QPushButton("Save Data")
        self.start_searching_button = QPushButton("Start Searching")
        # اتصال هر دکمه به یک تابع مربوطه
        self.list_of_file_button.clicked.connect(self.select_excel_file)
        self.source_pdf_address_button.clicked.connect(self.select_pdf_folder)
        self.save_data_button.clicked.connect(self.select_save_folder)
        self.start_searching_button.clicked.connect(self.start_searching)
        # ایجاد یک برچسب برای نمایش آدرس فایل اکسل
        self.excel_file_label = QLabel("No file selected")
        # تنظیم قابلیت شکستن خط برچسب
        self.excel_file_label.setWordWrap(True)
        # تنظیم حالت متن برچسب به وسط
        self.excel_file_label.setAlignment(Qt.AlignCenter)
        # ایجاد یک برچسب برای نمایش آدرس پوشه PDF
        self.pdf_folder_label = QLabel("No folder selected")
        # تنظیم قابلیت شکستن خط برچسب
        self.pdf_folder_label.setWordWrap(True)
        # تنظیم حالت متن برچسب به وسط
        self.pdf_folder_label.setAlignment(Qt.AlignCenter)
        # ایجاد یک برچسب برای نمایش آدرس پوشه ذخیره‌سازی
        self.save_folder_label = QLabel("No folder selected")
        # تنظیم قابلیت شکستن خط برچسب
        self.save_folder_label.setWordWrap(True)
        # تنظیم حالت متن برچسب به وسط
        self.save_folder_label.setAlignment(Qt.AlignCenter)
        # ایجاد یک نوار پیشرفت برای نمایش وضعیت جستجو
        self.progress_bar = QProgressBar()
        # تنظیم مقدار اولیه نوار پیشرفت به صفر
        self.progress_bar.setValue(0)
        # ایجاد یک چیدمان شبکه‌ای برای قرار دادن عناصر گرافیکی
        layout = QGridLayout()
        # اضافه کردن دکمه‌ها و برچسب‌ها به چیدمان با مشخص کردن سطر و ستون آن‌ها
        layout.addWidget(self.list_of_file_button, 0, 0)
        layout.addWidget(self.excel_file_label, 0, 1)
        layout.addWidget(self.source_pdf_address_button, 1, 0)
        layout.addWidget(self.pdf_folder_label, 1, 1)
        layout.addWidget(self.save_data_button, 2, 0)
        layout.addWidget(self.save_folder_label, 2, 1)
        layout.addWidget(self.start_searching_button, 3, 0)
        layout.addWidget(self.progress_bar, 3, 1)
        # تنظیم چیدمان به عنوان چیدمان پنجره
        self.setLayout(layout)
        # نمایش پنجره
        self.show()

    # تعریف یک تابع برای انتخاب فایل اکسل
    def select_excel_file(self):
        # باز کردن یک پنجره انتخاب فایل و دریافت آدرس فایل انتخاب شده
        file_name, _ = QFileDialog.getOpenFileName(self, "Select Excel File", "", "Excel Files (*.xlsx)")
        # اگر فایلی انتخاب شده باشد
        if file_name:
            # ذخیره آدرس فایل در یک متغیر کلاس
            self.excel_file = file_name
            # نمایش آدرس فایل در برچسب مربوطه
            self.excel_file_label.setText(file_name)

    # تعریف یک تابع برای انتخاب پوشه PDF
    def select_pdf_folder(self):
        # باز کردن یک پنجره انتخاب پوشه و دریافت آدرس پوشه انتخاب شده
        folder_name = QFileDialog.getExistingDirectory(self, "Select PDF Folder")
        # اگر پوشه‌ای انتخاب شده باشد
        if folder_name:
            # ذخیره آدرس پوشه در یک متغیر کلاس
            self.pdf_folder = folder_name
            # نمایش آدرس پوشه در برچسب مربوطه
            self.pdf_folder_label.setText(folder_name)

    # تعریف یک تابع برای انتخاب پوشه ذخیره‌سازی
    def select_save_folder(self):
        # باز کردن یک پنجره انتخاب پوشه و دریافت آدرس پوشه انتخاب شده
        folder_name = QFileDialog.getExistingDirectory(self, "Select Save Folder")
        # اگر پوشه‌ای انتخاب شده باشد
        if folder_name:
            # ذخیره آدرس پوشه در یک متغیر کلاس
            self.save_folder = folder_name
            # نمایش آدرس پوشه در برچسب مربوطه
            self.save_folder_label.setText(folder_name)

    # تعریف یک تابع برای شروع جستجو
    def start_searching(self):
        # بررسی اینکه آیا فایل اکسل، پوشه PDF و پوشه ذخیره‌سازی انتخاب شده‌اند
        if hasattr(self, "excel_file") and hasattr(self, "pdf_folder") and hasattr(self, "save_folder"):
            # ایجاد یک شی از کلاس SearchThread با مقادیر مورد نیاز
            self.search_thread = SearchThread(self.excel_file, self.pdf_folder, self.save_folder)
            # اتصال سیگنال پیشرفت جستجو به تابعی که مقدار نوار پیشرفت را تغییر می‌دهد
            self.search_thread.progress.connect(self.update_progress_bar)
            # شروع رشته جستجو
            self.search_thread.start()
        else:
            # نمایش یک پیام خطا به کاربر
            self.show_error("Please select the excel file, the pdf folder and the save folder first.")

    # تعریف یک تابع برای به‌روزرسانی مقدار نوار پیشرفت
    def update_progress_bar(self, value):
        # تنظیم مقدار نوار پیشرفت به مقدار دریافتی
        self.progress_bar.setValue(value)

    # تعریف یک تابع برای نمایش پیام خطا
    def show_error(self, message):
        # ایجاد یک برچسب با متن پیام خطا
        error_label = QLabel(message)
        # تنظیم رنگ متن برچسب به قرمز
        error_label.setStyleSheet("color: red;")
        # تنظیم حالت متن برچسب به وسط
        error_label.setAlignment(Qt.AlignCenter)
        # ایجاد یک پنجره جدید با عنوان Error
        error_window = QWidget()
        error_window.setWindowTitle("Error")
        # ایجاد یک چیدمان برای پنجره خطا
        error_layout = QGridLayout()
        # اضافه کردن برچسب خطا به چیدمان
        error_layout.addWidget(error_label)
        # تنظیم چیدمان به عنوان چیدمان پنجره خطا
        error_window.setLayout(error_layout)
        # نمایش پنجره خطا
        error_window.show()

# ایجاد یک شی از کلاس QApplication
app = QApplication(sys.argv)
# ایجاد یک شی از کلاس MainWindow
window = MainWindow()
# اجرای برنامه
sys.exit(app.exec_())
