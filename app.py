import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import openpyxl
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter
import os

class InventoryApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Aplikasi Inventarisasi Barang")
        self.root.geometry("800x600")
        self.root.resizable(True, True)
        
        # Nama file Excel
        self.filename = "DBR.xlsx"
        
        # Frame utama
        main_frame = ttk.Frame(root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Konfigurasi grid
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Judul
        title_label = ttk.Label(main_frame, text="FORM INPUT BARANG INVENTARIS", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20))
        
        # Form input
        self.create_input_form(main_frame)
        
        # Tombol aksi
        self.create_action_buttons(main_frame)
        
        # Preview data
        self.create_preview_section(main_frame)
        
    def create_input_form(self, parent):
        # Frame form
        form_frame = ttk.LabelFrame(parent, text="Input Data Barang", padding="10")
        form_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        form_frame.columnconfigure(1, weight=1)
        
        # Variabel untuk input
        self.no_urut_var = tk.StringVar()
        self.nama_barang_var = tk.StringVar()
        self.merk_type_var = tk.StringVar()
        self.kode_barang_var = tk.StringVar()
        self.tahun_perolehan_var = tk.StringVar()
        self.jumlah_var = tk.StringVar()
        self.keterangan_var = tk.StringVar()
        
        # Entri untuk Nomor Urut Pendaftaran
        ttk.Label(form_frame, text="Nomor Urut Pendaftaran:").grid(row=0, column=0, sticky=tk.W, pady=5)
        no_urut_entry = ttk.Entry(form_frame, textvariable=self.no_urut_var, width=30)
        no_urut_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=5, padx=(5, 0))
        
        # Entri untuk Nama Barang
        ttk.Label(form_frame, text="Nama Barang:").grid(row=1, column=0, sticky=tk.W, pady=5)
        nama_barang_entry = ttk.Entry(form_frame, textvariable=self.nama_barang_var, width=30)
        nama_barang_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=5, padx=(5, 0))
        
        # Entri untuk Merk/Type
        ttk.Label(form_frame, text="Merk/Type:").grid(row=2, column=0, sticky=tk.W, pady=5)
        merk_type_entry = ttk.Entry(form_frame, textvariable=self.merk_type_var, width=30)
        merk_type_entry.grid(row=2, column=1, sticky=(tk.W, tk.E), pady=5, padx=(5, 0))
        
        # Entri untuk Kode Barang
        ttk.Label(form_frame, text="Kode Barang:").grid(row=3, column=0, sticky=tk.W, pady=5)
        kode_barang_entry = ttk.Entry(form_frame, textvariable=self.kode_barang_var, width=30)
        kode_barang_entry.grid(row=3, column=1, sticky=(tk.W, tk.E), pady=5, padx=(5, 0))
        
        # Entri untuk Tahun Perolehan
        ttk.Label(form_frame, text="Tahun Perolehan:").grid(row=4, column=0, sticky=tk.W, pady=5)
        tahun_perolehan_entry = ttk.Entry(form_frame, textvariable=self.tahun_perolehan_var, width=30)
        tahun_perolehan_entry.grid(row=4, column=1, sticky=(tk.W, tk.E), pady=5, padx=(5, 0))
        
        # Entri untuk Jumlah Barang
        ttk.Label(form_frame, text="Jumlah Barang:").grid(row=5, column=0, sticky=tk.W, pady=5)
        jumlah_entry = ttk.Entry(form_frame, textvariable=self.jumlah_var, width=30)
        jumlah_entry.grid(row=5, column=1, sticky=(tk.W, tk.E), pady=5, padx=(5, 0))
        
        # Entri untuk Keterangan
        ttk.Label(form_frame, text="Keterangan:").grid(row=6, column=0, sticky=tk.W, pady=5)
        keterangan_entry = ttk.Entry(form_frame, textvariable=self.keterangan_var, width=30)
        keterangan_entry.grid(row=6, column=1, sticky=(tk.W, tk.E), pady=5, padx=(5, 0))
    
    def create_action_buttons(self, parent):
        # Frame tombol
        button_frame = ttk.Frame(parent)
        button_frame.grid(row=2, column=0, columnspan=2, pady=10)
        
        # Tombol Simpan
        save_button = ttk.Button(button_frame, text="Simpan Data", command=self.save_data)
        save_button.grid(row=0, column=0, padx=5)
        
        # Tombol Preview
        preview_button = ttk.Button(button_frame, text="Preview Template", command=self.preview_template)
        preview_button.grid(row=0, column=1, padx=5)
        
        # Tombol Cetak
        print_button = ttk.Button(button_frame, text="Cetak Template", command=self.print_template)
        print_button.grid(row=0, column=2, padx=5)
        
        # Tombol Reset
        reset_button = ttk.Button(button_frame, text="Reset Form", command=self.reset_form)
        reset_button.grid(row=0, column=3, padx=5)
    
    def create_preview_section(self, parent):
        # Frame preview
        preview_frame = ttk.LabelFrame(parent, text="Preview Data", padding="10")
        preview_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(10, 0))
        parent.rowconfigure(3, weight=1)
        preview_frame.columnconfigure(0, weight=1)
        preview_frame.rowconfigure(0, weight=1)
        
        # Treeview untuk preview
        columns = ("No", "No Urut", "Nama Barang", "Merk/Type", "Kode Barang", "Tahun", "Jumlah", "Keterangan")
        self.tree = ttk.Treeview(preview_frame, columns=columns, show="headings", height=10)
        
        # Mengatur heading
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100, anchor="center")
        
        # Scrollbar untuk treeview
        scrollbar = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        # Layout treeview dan scrollbar
        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Memuat data yang sudah ada
        self.load_existing_data()
    
    def is_merged_cell(self, sheet, row, col):
        """Cek apakah sel tertentu adalah bagian dari merged cells"""
        for merged_range in sheet.merged_cells.ranges:
            if merged_range.min_row <= row <= merged_range.max_row and \
               merged_range.min_col <= col <= merged_range.max_col:
                return True
        return False
    
    def get_next_available_row(self, sheet, start_row=20):
        """Cari baris kosong berikutnya yang tidak termasuk dalam merged cells"""
        row = start_row
        max_checks = 100  # Batasan untuk mencegah infinite loop
        
        for _ in range(max_checks):
            # Cek apakah sel A di baris ini kosong dan tidak tergabung
            if (sheet.cell(row=row, column=1).value is None and 
                not self.is_merged_cell(sheet, row, 1)):
                return row
            row += 1
        
        # Jika tidak ditemukan, kembalikan baris terakhir yang diperiksa
        return row
    
    def load_existing_data(self):
        # Hapus data lama di treeview
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Coba baca file Excel
        try:
            if os.path.exists(self.filename):
                workbook = openpyxl.load_workbook(self.filename)
                if "TEMPLATE" in workbook.sheetnames:
                    sheet = workbook["TEMPLATE"]
                    
                    # Cari baris data (mulai dari baris 20)
                    row_num = 20
                    data_found = False
                    
                    for row in sheet.iter_rows(min_row=20, max_row=100, values_only=True):
                        # Hanya proses baris yang memiliki data di kolom pertama
                        # dan bukan bagian dari merged cells
                        if row[0] is not None and str(row[0]).strip() and not self.is_merged_cell(sheet, row_num, 1):
                            try:
                                self.tree.insert("", "end", values=(
                                    row[0] or "",  # No
                                    row[1] or "",  # No Urut Pendaftaran
                                    row[2] or "",  # Nama Barang
                                    row[3] or "",  # Merk/Type
                                    row[4] or "",  # Kode Barang
                                    row[5] or "",  # Tahun Perolehan
                                    row[6] or "",  # Jumlah
                                    row[7] or ""   # Keterangan
                                ))
                                data_found = True
                            except Exception as e:
                                print(f"Error membaca baris {row_num}: {e}")
                        row_num += 1
                    
                    if not data_found:
                        print("Tidak ada data yang ditemukan atau semua data berada di merged cells")
                    
                    workbook.close()
        except Exception as e:
            messagebox.showerror("Error", f"Terjadi kesalahan saat membaca file: {str(e)}")
    
    def save_data(self):
        # Validasi input
        if not self.no_urut_var.get() or not self.nama_barang_var.get():
            messagebox.showerror("Error", "Nomor Urut Pendaftaran dan Nama Barang harus diisi!")
            return
        
        # Coba buka atau buat file Excel
        try:
            if os.path.exists(self.filename):
                workbook = openpyxl.load_workbook(self.filename)
            else:
                # Buat workbook baru dengan sheet TEMPLATE
                workbook = openpyxl.Workbook()
                # Hapus sheet default dan buat sheet TEMPLATE
                default_sheet = workbook.active
                workbook.remove(default_sheet)
                workbook.create_sheet("TEMPLATE")
            
            # Pastikan sheet TEMPLATE ada
            if "TEMPLATE" not in workbook.sheetnames:
                workbook.create_sheet("TEMPLATE")
            
            sheet = workbook["TEMPLATE"]
            
            # Cari baris kosong berikutnya yang aman (tidak di merged cells)
            row_num = self.get_next_available_row(sheet, 20)
            
            # Tulis data ke sheet
            try:
                sheet.cell(row=row_num, column=1, value=row_num-19)  # No. Urut
                sheet.cell(row=row_num, column=2, value=self.no_urut_var.get())  # Nomor Urut Pendaftaran
                sheet.cell(row=row_num, column=3, value=self.nama_barang_var.get())  # Nama Barang
                sheet.cell(row=row_num, column=4, value=self.merk_type_var.get())  # Merk/Type
                sheet.cell(row=row_num, column=5, value=self.kode_barang_var.get())  # Kode Barang
                sheet.cell(row=row_num, column=6, value=self.tahun_perolehan_var.get())  # Tahun Perolehan
                sheet.cell(row=row_num, column=7, value=self.jumlah_var.get())  # Jumlah
                sheet.cell(row=row_num, column=8, value=self.keterangan_var.get())  # Keterangan
                
                # Format template jika ini data pertama
                if row_num == 20:
                    self.format_template(sheet)
                
                # Simpan file
                workbook.save(self.filename)
                workbook.close()
                
                # Tampilkan pesan sukses
                messagebox.showinfo("Sukses", f"Data berhasil disimpan di baris {row_num}!")
                
                # Reset form dan muat ulang data
                self.reset_form()
                self.load_existing_data()
                
            except Exception as e:
                workbook.close()
                messagebox.showerror("Error", f"Terjadi kesalahan saat menulis data: {str(e)}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Terjadi kesalahan saat menyimpan data: {str(e)}")
    
    def format_template(self, sheet):
        """Format template hanya jika diperlukan dan hindari konflik dengan merged cells"""
        try:
            # Hapus merged cells yang mungkin mengganggu area data
            merged_ranges_to_remove = []
            for merged_range in list(sheet.merged_cells.ranges):
                if merged_range.min_row >= 20:  # Hapus merged cells di area data
                    merged_ranges_to_remove.append(merged_range)
            
            for merged_range in merged_ranges_to_remove:
                sheet.unmerge_cells(str(merged_range))
            
            # Format header template (area aman di atas baris 20)
            # Judul
            if sheet['A10'].value is None:
                sheet.merge_cells('A10:I12')
                sheet['A10'] = "DAFTAR BARANG RUANGAN (DBR)"
                sheet['A10'].alignment = Alignment(horizontal='center', vertical='center')
                sheet['A10'].font = Font(size=14, bold=True)
            
            # Informasi unit
            if sheet['A14'].value is None:
                sheet['A14'] = "Kode UAKPB"
                sheet['C14'] = ": 138.05.0200.693453.000KD"
                sheet['F14'] = "Ruangan"
                sheet['G14'] = ":"
                
                sheet['A15'] = "Nama Unit UAKPB"
                sheet['C15'] = ": BBPPMPV Pertanian Cianjur"
            
            # Header tabel
            if sheet['A17'].value is None:
                sheet.merge_cells('A17:I17')
                sheet['A17'] = "Nama Barang                 Tanda Pengenal Barang                                                             Keterangan"
                
                # Sub header
                sheet['A18'] = "No."
                sheet['B18'] = "Nomor Urut"
                sheet['C18'] = "Nama Barang"
                sheet['D18'] = "Merk/"
                sheet['E18'] = "Kode Barang"
                sheet['F18'] = "Tahun"
                sheet['G18'] = "Jumlah"
                sheet['H18'] = "Keterangan"
                
                sheet['A19'] = "Urut"
                sheet['B19'] = "Pendaftaran"
                sheet['D19'] = "Type"
                sheet['F19'] = "Perolehan"
                sheet['G19'] = "Barang"
            
            # Tanda tangan
            if sheet['A50'].value is None:
                sheet.merge_cells('A50:I50')
                sheet['A50'] = "Cianjur, November 2025"
                sheet['A50'].alignment = Alignment(horizontal='center')
                
                sheet['A52'] = "Penanggung Jawab UAKPB;"
                sheet['F52'] = "Penanggung Jawab Ruangan;"
                
                sheet['A53'] = "Kuasa Pengguna Barang"
                
                sheet['A55'] = "Dr. Yusuf, S.T., M.T."
                sheet['F55'] = "NIP. 196704181989121001"
                
                sheet['A56'] = "NIP. 196307201990011001"
                
        except Exception as e:
            print(f"Peringatan: Gagal memformat template: {e}")
    
    def preview_template(self):
        # Buka file Excel untuk preview
        try:
            if os.path.exists(self.filename):
                os.startfile(self.filename)  # Untuk Windows
            else:
                messagebox.showwarning("Peringatan", "File belum ada. Silakan simpan data terlebih dahulu.")
        except Exception as e:
            messagebox.showerror("Error", f"Tidak dapat membuka file: {str(e)}")
    
    def print_template(self):
        # Buka dialog print
        try:
            if os.path.exists(self.filename):
                os.startfile(self.filename, "print")  # Untuk Windows
                messagebox.showinfo("Info", "File telah dikirim ke printer. Pastikan printer siap.")
            else:
                messagebox.showwarning("Peringatan", "File belum ada. Silakan simpan data terlebih dahulu.")
        except Exception as e:
            messagebox.showerror("Error", f"Tidak dapat mencetak file: {str(e)}")
    
    def reset_form(self):
        # Reset semua field input
        self.no_urut_var.set("")
        self.nama_barang_var.set("")
        self.merk_type_var.set("")
        self.kode_barang_var.set("")
        self.tahun_perolehan_var.set("")
        self.jumlah_var.set("")
        self.keterangan_var.set("")

if __name__ == "__main__":
    root = tk.Tk()
    app = InventoryApp(root)
    root.mainloop()