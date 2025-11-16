import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Alignment, Font
import io
import base64
from datetime import datetime
import os

# Konfigurasi halaman
st.set_page_config(
    page_title="Aplikasi Inventarisasi Barang",
    page_icon="📦",
    layout="wide"
)

class InventoryAppStreamlit:
    def __init__(self):
        self.filename = "DBR.xlsx"
        self.initialize_session_state()
    
    def initialize_session_state(self):
        """Inisialisasi session state untuk menyimpan data sementara"""
        if 'data_barang' not in st.session_state:
            st.session_state.data_barang = []
        if 'edit_mode' not in st.session_state:
            st.session_state.edit_mode = False
        if 'edit_index' not in st.session_state:
            st.session_state.edit_index = None
    
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
        max_checks = 100
        
        for _ in range(max_checks):
            if (sheet.cell(row=row, column=1).value is None and 
                not self.is_merged_cell(sheet, row, 1)):
                return row
            row += 1
        return row
    
    def format_template(self, sheet):
        """Format template Excel"""
        try:
            # Hapus merged cells yang mengganggu area data
            merged_ranges_to_remove = []
            for merged_range in list(sheet.merged_cells.ranges):
                if merged_range.min_row >= 20:
                    merged_ranges_to_remove.append(merged_range)
            
            for merged_range in merged_ranges_to_remove:
                sheet.unmerge_cells(str(merged_range))
            
            # Format header template
            if sheet['A10'].value is None:
                sheet.merge_cells('A10:I12')
                sheet['A10'] = "DAFTAR BARANG RUANGAN (DBR)"
                sheet['A10'].alignment = Alignment(horizontal='center', vertical='center')
                sheet['A10'].font = Font(size=14, bold=True)
            
            if sheet['A14'].value is None:
                sheet['A14'] = "Kode UAKPB"
                sheet['C14'] = ": 138.05.0200.693453.000KD"
                sheet['F14'] = "Ruangan"
                sheet['G14'] = ":"
                
                sheet['A15'] = "Nama Unit UAKPB"
                sheet['C15'] = ": BBPPMPV Pertanian Cianjur"
            
            if sheet['A17'].value is None:
                sheet.merge_cells('A17:I17')
                sheet['A17'] = "Nama Barang                 Tanda Pengenal Barang                                                             Keterangan"
                
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
            
            if sheet['A50'].value is None:
                sheet.merge_cells('A50:I50')
                sheet['A50'] = f"Cianjur, {datetime.now().strftime('%B %Y')}"
                sheet['A50'].alignment = Alignment(horizontal='center')
                
                sheet['A52'] = "Penanggung Jawab UAKPB;"
                sheet['F52'] = "Penanggung Jawab Ruangan;"
                
                sheet['A53'] = "Kuasa Pengguna Barang"
                
                sheet['A55'] = "Dr. Yusuf, S.T., M.T."
                sheet['F55'] = "NIP. 196704181989121001"
                
                sheet['A56'] = "NIP. 196307201990011001"
                
        except Exception as e:
            st.warning(f"Peringatan dalam format template: {e}")
    
    def load_existing_data(self):
        """Muat data yang sudah ada dari file Excel"""
        try:
            if os.path.exists(self.filename):
                workbook = openpyxl.load_workbook(self.filename)
                if "TEMPLATE" in workbook.sheetnames:
                    sheet = workbook["TEMPLATE"]
                    
                    data = []
                    row_num = 20
                    
                    for row in sheet.iter_rows(min_row=20, max_row=100, values_only=True):
                        if (row[0] is not None and str(row[0]).strip() and 
                            not self.is_merged_cell(sheet, row_num, 1)):
                            data.append({
                                'No': row[0] or "",
                                'No Urut Pendaftaran': row[1] or "",
                                'Nama Barang': row[2] or "",
                                'Merk/Type': row[3] or "",
                                'Kode Barang': row[4] or "",
                                'Tahun Perolehan': row[5] or "",
                                'Jumlah': row[6] or "",
                                'Keterangan': row[7] or ""
                            })
                        row_num += 1
                    
                    workbook.close()
                    return data
            return []
        except Exception as e:
            st.error(f"Error membaca file: {e}")
            return []
    
    def save_to_excel(self, data):
        """Simpan data ke file Excel"""
        try:
            if os.path.exists(self.filename):
                workbook = openpyxl.load_workbook(self.filename)
            else:
                workbook = openpyxl.Workbook()
                default_sheet = workbook.active
                workbook.remove(default_sheet)
                workbook.create_sheet("TEMPLATE")
            
            if "TEMPLATE" not in workbook.sheetnames:
                workbook.create_sheet("TEMPLATE")
            
            sheet = workbook["TEMPLATE"]
            
            # Format template jika diperlukan
            self.format_template(sheet)
            
            # Cari baris kosong
            row_num = self.get_next_available_row(sheet, 20)
            
            # Tulis data
            sheet.cell(row=row_num, column=1, value=row_num-19)
            sheet.cell(row=row_num, column=2, value=data['no_urut'])
            sheet.cell(row=row_num, column=3, value=data['nama_barang'])
            sheet.cell(row=row_num, column=4, value=data['merk_type'])
            sheet.cell(row=row_num, column=5, value=data['kode_barang'])
            sheet.cell(row=row_num, column=6, value=data['tahun_perolehan'])
            sheet.cell(row=row_num, column=7, value=data['jumlah'])
            sheet.cell(row=row_num, column=8, value=data['keterangan'])
            
            workbook.save(self.filename)
            workbook.close()
            return True
            
        except Exception as e:
            st.error(f"Error menyimpan data: {e}")
            return False
    
    def get_excel_download_link(self):
        """Generate link untuk download file Excel"""
        try:
            if os.path.exists(self.filename):
                with open(self.filename, "rb") as file:
                    excel_data = file.read()
                
                b64 = base64.b64encode(excel_data).decode()
                href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{self.filename}">Download File Excel</a>'
                return href
            return None
        except Exception as e:
            st.error(f"Error generating download link: {e}")
            return None
    
    def render_form(self):
        """Render form input data"""
        st.header("📝 Form Input Data Barang")
        
        with st.form("input_form"):
            col1, col2 = st.columns(2)
            
            with col1:
                no_urut = st.text_input("Nomor Urut Pendaftaran*", 
                                      placeholder="Masukkan nomor urut pendaftaran")
                nama_barang = st.text_input("Nama Barang*", 
                                          placeholder="Masukkan nama barang")
                merk_type = st.text_input("Merk/Type", 
                                        placeholder="Masukkan merk/type")
                kode_barang = st.text_input("Kode Barang", 
                                          placeholder="Masukkan kode barang")
            
            with col2:
                tahun_perolehan = st.text_input("Tahun Perolehan", 
                                              placeholder="Masukkan tahun perolehan")
                jumlah = st.text_input("Jumlah Barang", 
                                     placeholder="Masukkan jumlah barang")
                keterangan = st.text_input("Keterangan", 
                                         placeholder="Masukkan keterangan")
            
            submitted = st.form_submit_button("💾 Simpan Data")
            
            if submitted:
                if not no_urut or not nama_barang:
                    st.error("❌ Nomor Urut Pendaftaran dan Nama Barang harus diisi!")
                else:
                    data = {
                        'no_urut': no_urut,
                        'nama_barang': nama_barang,
                        'merk_type': merk_type,
                        'kode_barang': kode_barang,
                        'tahun_perolehan': tahun_perolehan,
                        'jumlah': jumlah,
                        'keterangan': keterangan
                    }
                    
                    if self.save_to_excel(data):
                        st.success("✅ Data berhasil disimpan!")
                        st.session_state.data_barang = self.load_existing_data()
    
    def render_data_table(self):
        """Render tabel data barang"""
        st.header("📊 Data Barang Inventaris")
        
        # Muat data
        data = self.load_existing_data()
        
        if data:
            df = pd.DataFrame(data)
            st.dataframe(df, use_container_width=True)
            
            # Tombol aksi
            col1, col2, col3 = st.columns(3)
            
            with col1:
                if st.button("🔄 Refresh Data"):
                    st.session_state.data_barang = self.load_existing_data()
                    st.rerun()
            
            with col2:
                # Download link
                download_link = self.get_excel_download_link()
                if download_link:
                    st.markdown(download_link, unsafe_allow_html=True)
            
            with col3:
                if st.button("🖨️ Cetak/Export"):
                    self.export_to_excel()
        else:
            st.info("📭 Belum ada data barang. Silakan input data terlebih dahulu.")
    
    def export_to_excel(self):
        """Export data ke Excel untuk dicetak"""
        try:
            data = self.load_existing_data()
            if data:
                df = pd.DataFrame(data)
                
                # Buat file Excel dalam memory
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, sheet_name='DATA_BARANG', index=False)
                
                # Download link
                excel_data = output.getvalue()
                b64 = base64.b64encode(excel_data).decode()
                href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="data_barang_export.xlsx">📥 Download Data untuk Cetak</a>'
                st.markdown(href, unsafe_allow_html=True)
                
        except Exception as e:
            st.error(f"Error export data: {e}")
    
    def run(self):
        """Jalankan aplikasi"""
        st.title("📦 Aplikasi Inventarisasi Barang")
        st.markdown("---")
        
        # Tab untuk organisasi yang lebih baik
        tab1, tab2, tab3 = st.tabs(["Input Data", "Data Barang", "Informasi"])
        
        with tab1:
            self.render_form()
        
        with tab2:
            self.render_data_table()
        
        with tab3:
            st.header("ℹ️ Informasi Aplikasi")
            st.markdown("""
            ### Cara Penggunaan:
            1. **Input Data**: Gunakan tab "Input Data" untuk menambahkan barang inventaris
            2. **Lihat Data**: Gunakan tab "Data Barang" untuk melihat semua data yang tersimpan
            3. **Download**: Gunakan tombol download untuk mendapatkan file Excel
            
            ### Format File:
            - Data disimpan dalam file **DBR.xlsx**
            - Format mengikuti template DBR (Daftar Barangan Ruangan)
            - File dapat dibuka dengan Microsoft Excel atau aplikasi spreadsheet lainnya
            
            ### Kolom Wajib:
            - ❗ **Nomor Urut Pendaftaran**
            - ❗ **Nama Barang**
            """)

# Jalankan aplikasi
if __name__ == "__main__":
    app = InventoryAppStreamlit()
    app.run()
