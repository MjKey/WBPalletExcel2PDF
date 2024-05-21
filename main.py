import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from reportlab.lib.pagesizes import landscape, A4
from reportlab.pdfgen import canvas
from reportlab.lib import utils
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from barcode import Code128
from barcode.writer import ImageWriter
from tkcalendar import DateEntry
import os
import textwrap

class PalletApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel to PDF for Wildberries Pallet")

        self.destinations = ["Электросталь","Белые Столбы"]
        self.delivery_types = ["Монопалета"]
        self.company_names = ["ИП Иванов Иван Иванович", "ИП Дмитров Двитрий Дмитрович"]

        self.create_label_entry("Номер поставки", 0)
        self.create_dropdown("Склад назначения", self.destinations, 1)
        self.create_dropdown("Тип поставки", self.delivery_types, 2)
        self.create_dropdown("Наименование Юр. Лица", self.company_names, 3)
        self.create_date_entry("Дата поставки", 4)
        

        self.file_label = tk.Label(root, text="Выберите Excel файл с данными:")
        self.file_label.grid(row=5, column=0, padx=10, pady=5)

        self.file_button = tk.Button(root, text="Загрузить файл", command=self.load_file)
        self.file_button.grid(row=5, column=1, padx=10, pady=5)

        self.generate_button = tk.Button(root, text="Создать PDF для паллетов", command=self.generate_reports)
        self.generate_button.grid(row=6, columnspan=2, pady=20)

    def create_label_entry(self, label_text, row):
        label = tk.Label(self.root, text=label_text)
        label.grid(row=row, column=0, padx=10, pady=5)
        entry = tk.Entry(self.root)
        entry.grid(row=row, column=1, padx=10, pady=5)
        setattr(self, f"{label_text.replace(' ', '_').replace('.', '').lower()}_entry", entry)
    
    def create_dropdown(self, label_text, options, row):
        label = tk.Label(self.root, text=label_text)
        label.grid(row=row, column=0, padx=10, pady=5)
        selected_option = tk.StringVar(self.root)
        selected_option.set(options[0])
        dropdown = tk.OptionMenu(self.root, selected_option, *options)
        dropdown.grid(row=row, column=1, padx=10, pady=5)
        setattr(self, f"{label_text.replace(' ', '_').replace('.', '').lower()}_dropdown", selected_option)
    
    def create_date_entry(self, label_text, row):
        label = tk.Label(self.root, text=label_text)
        label.grid(row=row, column=0, padx=10, pady=5)
        date_entry = DateEntry(self.root, date_pattern="dd.mm.yyyy", locale="ru_RU")
        date_entry.grid(row=row, column=1, padx=10, pady=5)
        setattr(self, f"{label_text.replace(' ', '_').replace('.', '').lower()}_entry", date_entry)

    def load_file(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if self.file_path:
            self.file_label.config(text=f"Файл: {self.file_path}")

    def generate_reports(self):
        try:
            # Read user inputs
            delivery_number = self.номер_поставки_entry.get()
            destination = self.склад_назначения_dropdown.get()
            delivery_type = self.тип_поставки_dropdown.get()
            company_name = self.наименование_юр_лица_dropdown.get()
            delivery_date = self.дата_поставки_entry.get_date().strftime('%d.%m.%Y') 

            df = pd.read_excel(self.file_path)

            pallet_groups = df.groupby('Номер').agg({'Количество': 'sum', 'Код товара': lambda x: list(x), 'Штрихкод': lambda x: list(x)}).reset_index()

            for _, row in pallet_groups.iterrows():
                pallet_number = row['Номер']
                box_count = row['Количество']
                codes = row['Код товара']
                barcodes = row['Штрихкод']
                self.create_pdf_report(pallet_number, len(pallet_groups), box_count, delivery_number, destination, delivery_type, company_name, delivery_date, codes, barcodes)

            messagebox.showinfo("Успех", "PDF файлы успешно созданы!")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось создать PDF: {e}")

    def create_pdf_report(self, pallet_number, total_pallets, box_count, delivery_number, destination, delivery_type, company_name, delivery_date, codes, barcodes):
        directory = f"./{delivery_number}"
        if not os.path.exists(directory):
            os.makedirs(directory)
        file_name = f"./{directory}/{pallet_number}.pdf"
        c = canvas.Canvas(file_name, pagesize=landscape(A4))

        table_x = 50
        table_y = 550
        row_height = 48
        col_width = 200

        pdfmetrics.registerFont(TTFont('Calibri', 'calibri.ttf'))
        pdfmetrics.registerFont(TTFont('Calibri-Bold', 'calibrib.ttf'))
        c.setFont('Calibri', 20)

        data = [
            ("Номер палеты", pallet_number),
            ("Количество палет в поставке", total_pallets),
            ("Количество коробок на данной палете", box_count),
            ("Номер поставки", delivery_number),
            ("Склад назначения", destination),
            ("Тип поставки", delivery_type),
            ("Наименование Юр. Лица", company_name),
            ("Дата поставки", delivery_date)
        ]

        for i in range(len(data)):
            c.rect(table_x, table_y - (i + 1) * row_height, col_width, row_height)
            c.rect(table_x + col_width, table_y - (i + 1) * row_height, col_width + 5, row_height)

        for i, (label, value) in enumerate(data):
            wrapped_label = textwrap.fill(str(label), width=20)
            wrapped_value = textwrap.fill(str(value), width=20)

            text_label = wrapped_label.split('\n')
            text_value = wrapped_value.split('\n')

            for j, line in enumerate(text_label):
                c.drawString(table_x + 5, table_y - (i + 1) * row_height + row_height - 20 - (j * 20), line)

            for j, line in enumerate(text_value):
                c.drawString(table_x + col_width + 5, table_y - (i + 1) * row_height + row_height - 20 - (j * 20), line)

        right_x = table_x + 2 * col_width + 50
        barcode_y = table_y - 15
        c.drawString(right_x, barcode_y, "Баркоды:")
        barcode_y -= row_height

        for code in codes:
            c.drawString(right_x + 20, barcode_y + 20, str(code))
            barcode_y -= row_height // 2

        barcode_image = Code128(str(barcodes[0]), writer=ImageWriter())
        barcode_image.save(f"barcode_{pallet_number}")
        c.drawImage(f"barcode_{pallet_number}.png", right_x, barcode_y - 150, width=250, height=150)

        os.remove(f"barcode_{pallet_number}.png")

        c.save()

if __name__ == "__main__":
    root = tk.Tk()
    icon_path = "./1.ico"
    root.iconbitmap(icon_path)
    app = PalletApp(root)
    root.mainloop()
