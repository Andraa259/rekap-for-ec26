import streamlit as st
import pandas as pd
import io
import re
import plotly.express as px
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Wonderful Class Presence Hub", layout="wide")

st.title("📊 Wonderful Class Presence Dashboard & Fixed Matrix")
st.write("Sistem Matriks Presensi dengan Struktur Kolom Tetap (No, Nama, NBI, Total Hadir, P1 - P11).")

# ----------------- SIDEBAR INPUT -----------------
st.sidebar.header("📁 Unggah Dokumen Presensi")

master_file = st.sidebar.file_uploader(
    "1. File Master Database Anggota (Harus ada kolom 'NBI' & 'Nama')", 
    type=["xlsx", "xls"],
    key="master_upload"
)

source_files = st.sidebar.file_uploader(
    "2. Unggah Form Absensi Harian (Pertemuan 1 - 11)", 
    type=["xlsx", "xls"], 
    accept_multiple_files=True,
    key="sources_upload"
)

# ----------------- LOGIKA UTAMA -----------------
if st.sidebar.button("🚀 Proses & Sinkronisasi Presensi", type="primary"):
    if not master_file:
        st.error("❌ Mohon unggah File Master Database Anggota terlebih dahulu!")
    elif not source_files:
        st.error("❌ Mohon unggah minimal satu File Form Absensi Harian!")
    else:
        with st.spinner("Sedang menyusun matriks kehadiran sesuai format..."):
            try:
                # 1. Baca Master & Standarisasi Data Awal
                df_master_raw = pd.read_excel(master_file)
                df_master_raw.columns = df_master_raw.columns.str.strip()
                
                if 'NBI' not in df_master_raw.columns or 'Nama' not in df_master_raw.columns:
                    st.error("❌ File Master harus memiliki kolom bernama 'NBI' dan 'Nama'!")
                    st.stop()
                
                df_master_raw['NBI'] = df_master_raw['NBI'].astype(str).str.strip()
                df_master_raw['Nama'] = df_master_raw['Nama'].astype(str).str.strip()
                
                # Buat DataFrame Baru dengan Struktur Kolom Sesuai Request
                df_rekap = pd.DataFrame()
                df_rekap['No.'] = range(1, len(df_master_raw) + 1)
                df_rekap['Nama'] = df_master_raw['Nama']
                df_rekap['NBI'] = df_master_raw['NBI']
                df_rekap['Total Hadir'] = 0  # Inisialisasi awal
                
                # Buat kolom Pertemuan 1 sampai Pertemuan 11 (Default: Alpa)
                list_pertemuan = [f"Pertemuan {i}" for i in range(1, 12)]
                for p in list_pertemuan:
                    df_rekap[p] = "Alpa"

                # 2. Proses File Form Absensi yang Diunggah
                kehadiran_per_hari = {}
                
                for src in source_files:
                    df_src = pd.read_excel(src)
                    df_src.columns = df_src.columns.str.strip()
                    
                    if 'NBI' not in df_src.columns:
                        st.warning(f"⚠️ File '{src.name}' dilewati karena tidak memiliki kolom 'NBI'.")
                        continue
                        
                    df_src['NBI'] = df_src['NBI'].astype(str).str.strip()
                    
                    # ANTI CURANG: Ambil data pertama jika absen ganda
                    df_src = df_src.drop_duplicates(subset=['NBI'], keep='first')
                    
                    # Deteksi Angka Pertemuan dari Nama File (Misal: "Absen P1.xlsx" atau "Form 2" -> mendeteksi angka)
                    nama_file = src.name.lower()
                    angka_terdeteksi = re.findall(r'\d+', nama_file)
                    
                    if angka_terdeteksi:
                        nomor_p = int(angka_terdeteksi[0])
                        if 1 <= nomor_p <= 11:
                            target_kolom = f"Pertemuan {nomor_p}"
                            
                            # Tandai Hadir bagi yang NBI-nya terdaftar di file form ini
                            df_rekap[target_kolom] = df_rekap['NBI'].isin(df_src['NBI']).map({True: 'Hadir', False: 'Alpa'})
                            
                            # Hitung statistik untuk grafik batang
                            total_hadir = (df_rekap[target_kolom] == 'Hadir').sum()
                            kehadiran_per_hari[target_kolom] = total_hadir
                        else:
                            st.warning(f"⚠️ File '{src.name}' memiliki angka pertemuan di luar jangkauan 1-11.")
                    else:
                        st.warning(f"⚠️ Gagal mendeteksi angka pertemuan pada nama file '{src.name}'. Pastikan ada angka 1-11 di nama filenya.")

                # 3. Hitung Total Hadir Secara Aktual untuk Tiap Baris
                df_rekap['Total Hadir'] = (df_rekap[list_pertemuan] == 'Hadir').sum(axis=1)

                # ----------------- VISUALISASI DASHBOARD -----------------
                st.subheader("📌 Ringkasan Metrik Kelas")
                m1, m2, m3 = st.columns(3)
                m1.metric("Total Anggota Kelas", f"{len(df_rekap)} Orang")
                m2.metric("Total Kolom Pertemuan", "11 Sesi")
                
                total_maksimal_hadir = len(df_rekap) * 11
                persentase_total = ((df_rekap['Total Hadir'].sum() / total_maksimal_hadir) * 100).round(1)
                m3.metric("Rata-rata Presensi Kelas Overall", f"{persentase_total} %")
                
                if kehadiran_per_hari:
                    st.markdown("---")
                    st.subheader("📈 Visualisasi Grafik Batang Kehadiran Per Sesi")
                    # Urutkan kunci agar grafik tampil berurutan dari P1 ke P11
                    sorted_kehadiran = sorted(kehadiran_per_hari.items(), key=lambda x: int(re.findall(r'\d+', x[0])[0]))
                    df_grafik = pd.DataFrame(sorted_kehadiran, columns=['Pertemuan', 'Jumlah Hadir'])
                    
                    fig = px.bar(
                        df_grafik, x='Pertemuan', y='Jumlah Hadir', text='Jumlah Hadir',
                        color_discrete_sequence=['#1E3A8A']
                    )
                    fig.update_traces(textposition='outside')
                    st.plotly_chart(fig, use_container_width=True)

                st.markdown("---")
                st.subheader("👀 Preview Tabel Matriks Terformat")
                st.dataframe(df_rekap, use_container_width=True)

                # --- EXCEL FORMATTING WITH OPENPYXL (STRIKTUR KETat) ---
                wb = Workbook()
                ws = wb.active
                ws.title = "Rekap Presensi Wonderful"
                ws.views.sheetView[0].showGridLines = True

                columns_list = df_rekap.columns.tolist()
                total_cols = len(columns_list)
                
                # Kolom B adalah kolom pertama data, Kolom terakhir bergeser karena Kolom A kosong
                last_col_letter = get_column_letter(1 + total_cols)

                # 1. Judul Utama di Baris 2 & 3 (Merge dari B sampai Kolom Akhir, Center)
                ws.merge_cells(f"B2:{last_col_letter}2")
                ws.merge_cells(f"B3:{last_col_letter}3")
                ws["B2"] = "REKAPITULASI PRESENSI"
                ws["B3"] = "WONDERFUL CLASS 2026"

                title_font = Font(name="Calibri", size=14, bold=True)
                center_alignment = Alignment(horizontal="center", vertical="center")
                
                ws["B2"].font = title_font
                ws["B2"].alignment = center_alignment
                ws["B3"].font = title_font
                ws["B3"].alignment = center_alignment
                
                ws.row_dimensions[2].height = 24
                ws.row_dimensions[3].height = 24

                # 2. Baris Judul Tabel / Header (Baris 6) -> Bold Teks & Bold/Medium Border
                header_row = 6
                ws.row_dimensions[header_row].height = 28
                
                thin_side = Side(border_style="thin", color="000000")
                thick_side = Side(border_style="medium", color="000000")
                
                data_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
                header_border = Border(left=thin_side, right=thin_side, top=thick_side, bottom=thick_side)
                
                header_font = Font(name="Calibri", size=11, bold=True)
                header_fill = PatternFill(start_color="EAEAEA", end_color="EAEAEA", fill_type="solid")

                for col_idx, col_name in enumerate(columns_list, start=2):
                    cell = ws.cell(row=header_row, column=col_idx)
                    cell.value = col_name
                    cell.font = header_font
                    cell.fill = header_fill
                    cell.alignment = center_alignment
                    cell.border = header_border

                # 3. Pengisian Data & Semua Cell Harus Center & Border (Baris 7 Dst)
                data_font = Font(name="Calibri", size=11)
                current_row = 7
                
                for _, row_data in df_rekap.iterrows():
                    ws.row_dimensions[current_row].height = 20
                    for col_idx, col_name in enumerate(columns_list, start=2):
                        cell = ws.cell(row=current_row, column=col_idx)
                        val = row_data[col_name]
                        cell.value = "" if pd.isna(val) else val
                        cell.font = data_font
                        cell.alignment = center_alignment
                        cell.border = data_border
                    current_row += 1

                # 4. Pengaturan Lebar Kolom (Kolom A Sengaja Kosong)
                ws.column_dimensions['A'].width = 4
                for col_idx in range(2, 2 + total_cols):
                    col_letter = get_column_letter(col_idx)
                    max_len = max(len(str(ws.cell(row=r, column=col_idx).value or '')) for r in range(6, current_row))
                    ws.column_dimensions[col_letter].width = max(max_len + 5, 12)

                # Export hasil ke memory untuk di-download
                output = io.BytesIO()
                wb.save(output)
                processed_data = output.getvalue()

                st.sidebar.markdown("---")
                st.sidebar.success("🎉 Matriks Presensi Berhasil Disusun!")
                st.sidebar.download_button(
                    label="📥 Download Excel Hasil Rekap Fixed",
                    data=processed_data,
                    file_name="Rekap_Fixed_WonderfulClass_2026.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"Terjadi kesalahan teknis: {str(e)}")
