import tkinter as tk
from tkinter import filedialog, messagebox
import fitz  # PyMuPDF
import csv

class PDFtoCSVConverter(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PDF to CSV Converter")
        self.geometry("500x350")

        tk.Label(self, text="Target HEX Color:").pack()
        self.color_entry = tk.Entry(self, width=50)
        self.color_entry.insert(0, "#005880")  # Default color
        self.color_entry.pack()

        tk.Label(self, text="Tolerance:").pack()
        self.tolerance_entry = tk.Entry(self, width=50)
        self.tolerance_entry.insert(0, "15")  # Default tolerance
        self.tolerance_entry.pack()

        tk.Label(self, text="Select PDF:").pack()
        self.pdf_entry = tk.Entry(self, width=50)
        self.pdf_entry.pack()
        tk.Button(self, text="Browse...", command=self.select_pdf).pack()

        tk.Label(self, text="Save CSV As:").pack()
        self.csv_entry = tk.Entry(self, width=50)
        self.csv_entry.pack()
        tk.Button(self, text="Browse...", command=self.select_csv).pack()

        tk.Label(self, text="Skip Entries (comma-separated):").pack()
        self.skip_entry = tk.Entry(self, width=50)
        self.skip_entry.pack()

        tk.Button(self, text="Convert", command=self.convert).pack()

    def select_pdf(self):
        file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        self.pdf_entry.delete(0, tk.END)
        self.pdf_entry.insert(0, file_path)

    def select_csv(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
        self.csv_entry.delete(0, tk.END)
        self.csv_entry.insert(0, file_path)

    def convert(self):
        pdf_path = self.pdf_entry.get()
        csv_path = self.csv_entry.get()
        target_color_hex = self.color_entry.get().strip()
        tolerance = int(self.tolerance_entry.get().strip())
        skip_sentences = [sentence.strip() for sentence in self.skip_entry.get().split(',')]

        if not pdf_path.lower().endswith('.pdf'):
            messagebox.showerror("Error", "Selected file is not a PDF. Please select a PDF file.")
            return
        
        try:
            data_for_csv = self.extract_titles_and_page_numbers(pdf_path, target_color_hex, tolerance, skip_sentences)
            self.save_csv(data_for_csv, csv_path)
            messagebox.showinfo("Success", "Conversion complete! CSV saved.")
        except Exception as e:
            messagebox.showerror("Failure", f"Conversion failed: {e}")

    @staticmethod
    def hex_to_rgb(hex_color):
        hex_color = hex_color.lstrip('#')
        return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))

    @staticmethod
    def int_to_rgb(int_color):
        int_color = int(int_color)  # Ensure int_color is an integer
        return ((int_color >> 16) & 255, (int_color >> 8) & 255, int_color & 255)


    @staticmethod
    def is_color_within_range(target_color, actual_color, tolerance):
        return all(abs(tc - ac) <= tolerance for tc, ac in zip(target_color, actual_color))

    def extract_titles_and_page_numbers(self, pdf_path, target_color_hex, tolerance, skip_sentences):
        target_color_rgb = self.hex_to_rgb(target_color_hex)
        doc = fitz.open(pdf_path)
        data_for_csv = []
        for page_num, page in enumerate(doc):
            text_instances = page.get_text("words")
            for instance in text_instances:
                text = instance[4]
                color = instance[3]
                actual_color_rgb = self.int_to_rgb(color)
                if self.is_color_within_range(target_color_rgb, actual_color_rgb, tolerance) and not any(skip_sentence.lower() in text.lower() for skip_sentence in skip_sentences):
                    data_for_csv.append((text, str(page_num + 1)))
        doc.close()
        return data_for_csv

    def save_csv(self, data, csv_path):
        with open(csv_path, 'w', newline='', encoding='utf-8-sig') as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow(['Text', 'Page Number'])
            for row in data:
                writer.writerow(row)

if __name__ == "__main__":
    app = PDFtoCSVConverter()
    app.mainloop()
