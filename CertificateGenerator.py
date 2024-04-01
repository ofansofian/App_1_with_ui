import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image, ImageDraw, ImageFont
import pandas as pd
import os
from reportlab.lib.pagesizes import letter, landscape, portrait
from reportlab.pdfgen import canvas

class CertificateGenerator:
    def __init__(self, master):
        self.master = master
        self.master.title("NFF E-Certificate Generator")

        self.template_path = tk.StringVar()
        self.font_path = tk.StringVar()
        self.data_path = tk.StringVar()
        self.output_folder = tk.StringVar()
        self.nama_x = tk.DoubleVar()
        self.nama_y = tk.DoubleVar()
        self.font_size = tk.IntVar(value=30)
        self.orientation = tk.StringVar(value="portrait")

        self.create_widgets()

    def create_widgets(self):
        tk.Label(self.master, text="NFF E-Certificate Generator", font=("Helvetica", 16, "bold")).pack(pady=10)

        input_frame = tk.Frame(self.master)
        input_frame.pack(pady=10)

        tk.Label(input_frame, text="Template Sertifikat:").grid(row=0, column=0, padx=5, pady=5)
        self.template_entry = tk.Entry(input_frame, width=50, textvariable=self.template_path)
        self.template_entry.grid(row=0, column=1, padx=5, pady=5)
        tk.Button(input_frame, text="Browse", command=self.browse_template).grid(row=0, column=2, padx=5, pady=5)

        tk.Label(input_frame, text="Font:").grid(row=1, column=0, padx=5, pady=5)
        self.font_entry = tk.Entry(input_frame, width=50, textvariable=self.font_path)
        self.font_entry.grid(row=1, column=1, padx=5, pady=5)
        tk.Button(input_frame, text="Browse", command=self.browse_font).grid(row=1, column=2, padx=5, pady=5)

        tk.Label(input_frame, text="Data Peserta:").grid(row=2, column=0, padx=5, pady=5)
        self.data_entry = tk.Entry(input_frame, width=50, textvariable=self.data_path)
        self.data_entry.grid(row=2, column=1, padx=5, pady=5)
        tk.Button(input_frame, text="Browse", command=self.browse_data).grid(row=2, column=2, padx=5, pady=5)

        tk.Label(input_frame, text="Output Folder:").grid(row=3, column=0, padx=5, pady=5)
        self.output_entry = tk.Entry(input_frame, width=50, textvariable=self.output_folder)
        self.output_entry.grid(row=3, column=1, padx=5, pady=5)
        tk.Button(input_frame, text="Browse", command=self.browse_output_folder).grid(row=3, column=2, padx=5, pady=5)

        tk.Label(input_frame, text="Nama X:").grid(row=4, column=0, padx=5, pady=5)
        self.nama_x_entry = tk.Entry(input_frame, textvariable=self.nama_x)
        self.nama_x_entry.grid(row=4, column=1, padx=5, pady=5)

        tk.Label(input_frame, text="Nama Y:").grid(row=5, column=0, padx=5, pady=5)
        self.nama_y_entry = tk.Entry(input_frame, textvariable=self.nama_y)
        self.nama_y_entry.grid(row=5, column=1, padx=5, pady=5)

        tk.Label(input_frame, text="Font Size:").grid(row=6, column=0, padx=5, pady=5)
        self.font_size_entry = tk.Entry(input_frame, textvariable=self.font_size)
        self.font_size_entry.grid(row=6, column=1, padx=5, pady=5)

        orientation_frame = tk.LabelFrame(self.master, text="Layout")
        orientation_frame.pack(pady=10)
        ttk.Radiobutton(orientation_frame, text="Potrait", variable=self.orientation, value="portrait").grid(row=0, column=0, padx=5, pady=5)
        ttk.Radiobutton(orientation_frame, text="Landscape", variable=self.orientation, value="landscape").grid(row=0, column=1, padx=5, pady=5)

        tk.Button(self.master, text="Preview Sertifikat", command=self.preview_certificates).pack(pady=10)
        tk.Button(self.master, text="Generate Sertifikat", command=self.generate_certificates).pack(pady=5)

    def browse_template(self):
        filename = filedialog.askopenfilename()
        self.template_entry.delete(0, tk.END)
        self.template_entry.insert(0, filename)

    def browse_font(self):
        filename = filedialog.askopenfilename()
        self.font_entry.delete(0, tk.END)
        self.font_entry.insert(0, filename)

    def browse_data(self):
        filename = filedialog.askopenfilename()
        self.data_entry.delete(0, tk.END)
        self.data_entry.insert(0, filename)

    def browse_output_folder(self):
        foldername = filedialog.askdirectory()
        self.output_entry.delete(0, tk.END)
        self.output_entry.insert(0, foldername)

    def preview_certificates(self):
        try:
            template_path = self.template_path.get()
            font_path = self.font_path.get()
            data_path = self.data_path.get()
            font_size = self.font_size.get()

            template = Image.open(template_path)
            font = ImageFont.truetype(font_path, size=font_size)
            data = pd.read_excel(data_path)

            for index, row in data.iterrows():
                sertifikat = template.copy()
                draw = ImageDraw.Draw(sertifikat)
                nama = row['Nama']
                posisi_nama = (self.nama_x.get(), self.nama_y.get())

                draw.text(posisi_nama, nama, fill="black", font=font)

                sertifikat.show()
        except Exception as e:
            messagebox.showerror("Error", f"Terjadi kesalahan: {str(e)}")

    def generate_certificates(self):
        try:
            template_path = self.template_path.get()
            font_path = self.font_path.get()
            data_path = self.data_path.get()
            output_folder = self.output_folder.get()
            font_size = self.font_size.get()
            orientation = self.orientation.get()

            if orientation == "portrait":
                page_size = portrait(letter)
            else:
                page_size = landscape(letter)

            template = Image.open(template_path)
            font = ImageFont.truetype(font_path, size=font_size)
            data = pd.read_excel(data_path)

            pdf_path = os.path.join(output_folder, "sertifikat.pdf")
            c = canvas.Canvas(pdf_path, pagesize=page_size)

            for index, row in data.iterrows():
                nama = row['Nama']
                posisi_nama = (self.nama_x.get(), self.nama_y.get())

                sertifikat = template.copy()
                draw = ImageDraw.Draw(sertifikat)
                draw.text(posisi_nama, nama, fill="black", font=font)

                png_path = os.path.join(output_folder, f"sertifikat_{nama}.png")
                sertifikat.save(png_path)

                c.drawImage(png_path, 0, 0, width=page_size[0], height=page_size[1])
                c.showPage()

                os.remove(png_path)

            c.save()

            messagebox.showinfo("Success", "Sertifikat berhasil dibuat!")
        except Exception as e:
            messagebox.showerror("Error", f"Terjadi kesalahan: {str(e)}")

def main():
    root = tk.Tk()
    app = CertificateGenerator(root)
    root.mainloop()

if __name__ == "__main__":
    main()
