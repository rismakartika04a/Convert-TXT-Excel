import streamlit as st
import re
import os
import tempfile
import zipfile
import pandas as pd
import numpy as np

st.set_page_config(layout="wide")
st.title("Split file TXT & Convert ➔ Excel")

uploaded_file = st.file_uploader("Upload file TXT", type=None)
if uploaded_file is None:
    st.info("Silakan upload file TXT...")
    st.stop()

split_header = st.selectbox(
    "Pilih header pemisah halaman (opsional):",
    ["", "GAJI KARYAWAN TETAP", "GAJI KARYAWAN KONTRAK", "LAPORAN ABSENSI STAFF"]
)

# --- Baca isi file
try:
    try:
        raw_text = uploaded_file.read().decode("utf-8")
    except UnicodeDecodeError:
        uploaded_file.seek(0)
        raw_text = uploaded_file.read().decode("latin-1")
except Exception as e:
    st.error(f"Gagal membaca file: {e}")
    st.stop()

# --- Split berdasarkan header
if not split_header:
    sections = [raw_text]
    st.info("Tidak ada header dipilih → 1 section saja.")
else:
    pattern = rf'({re.escape(split_header)}.*?)(?={re.escape(split_header)}|\Z)'
    sections = re.findall(pattern, raw_text, flags=re.DOTALL)
    if not sections:
        sections = [raw_text]
        st.warning("Header tidak ditemukan, semua jadi 1 section.")
    else:
        st.success(f"Ditemukan {len(sections)} section berdasarkan header.")

# Format preview angka Indonesia
def format_id(x):
    if pd.isna(x) or x == "":
        return ""
    try:
        return "{:,.0f}".format(float(x)).replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return x

def get_preview_df(df):
    df_preview = df.copy()
    for col in df_preview.columns:
        if pd.api.types.is_numeric_dtype(df_preview[col]):
            df_preview[col] = df_preview[col].apply(format_id)
    return df_preview

def is_mostly_number(parts):
    count_num = 0
    count_nonnum = 0
    for x in parts:
        x = x.replace(",", "").replace(".", "").replace(" ", "")
        if x == "": continue
        if x.replace("-", "").isdigit():
            count_num += 1
        else:
            count_nonnum += 1
    if count_num == 0: return False
    return count_num >= max(2, int(0.5 * (count_num + count_nonnum)))

def detect_header_lines(lines, delimiter="³", min_header=2, max_header=4):
    for i, ln in enumerate(lines):
        if "NIK" in ln.upper() and "NAMA" in ln.upper():
            header_lines = []
            for j in range(max_header):
                idx = i + j
                if idx < len(lines) and delimiter in lines[idx]:
                    parts = [p.strip() for p in lines[idx].split(delimiter)]
                    if is_mostly_number(parts): break
                    header_lines.append(lines[idx])
            return i, header_lines
    for i, ln in enumerate(lines):
        if ln.count(delimiter) >= min_header:
            header_lines = []
            for j in range(max_header):
                idx = i + j
                if idx < len(lines) and delimiter in lines[idx]:
                    parts = [p.strip() for p in lines[idx].split(delimiter)]
                    if is_mostly_number(parts): break
                    header_lines.append(lines[idx])
            return i, header_lines
    return 0, []

# Fungsi membuat nama kolom unik
def make_columns_unique(columns):
    counts = {}
    new_cols = []
    for col in columns:
        if col == "":
            col = "UNNAMED"
        key = col.strip()
        if key in counts:
            counts[key] += 1
            new_cols.append(f"{key}_{counts[key]}")
        else:
            counts[key] = 1
            new_cols.append(key)
    return new_cols

with tempfile.TemporaryDirectory() as tmpdirname:
    zip_path = os.path.join(tmpdirname, "hasil_split_gaji_karyawan_txt.zip")
    excel_path = os.path.join(tmpdirname, "hasil_split_gaji_karyawan.xlsx")
    preview_dfs = []
    filepaths = []

    # Simpan zip dulu (TXT-TXT)
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zipf:
        # Tulis semua excel sheet dulu ke writer
        with pd.ExcelWriter(excel_path, engine="xlsxwriter") as writer:
            for idx, section in enumerate(sections, start=1):
                fn_txt = f"gaji_karyawan_page_{idx}.txt"
                fp_txt = os.path.join(tmpdirname, fn_txt)
                with open(fp_txt, "w", encoding="utf-8") as f:
                    f.write(section)
                zipf.write(fp_txt, arcname=fn_txt)
                filepaths.append((fn_txt, fp_txt, section))

                delimiter = "³"
                lines = [ln for ln in section.splitlines() if delimiter in ln]
                if not lines:
                    st.warning(f"Section #{idx}: Tidak ada data (delimiter ³).")
                    df = pd.DataFrame()
                else:
                    # PATCH HEADER
                    header_idx, header_lines = detect_header_lines(lines, delimiter=delimiter, min_header=2, max_header=4)
                    if not header_lines:
                        header_lines = [lines[0]]

                    # Bikin parts header per baris
                    header_parts = [ [p.strip() for p in h.replace("Â", "").replace("Ã", "").replace("°", "").split(delimiter)] for h in header_lines ]
                    maxlen = max(len(x) for x in header_parts)
                    for parts in header_parts:
                        while len(parts) < maxlen:
                            parts.append("")

                    # Gabung header per kolom
                    columns = []
                    for i in range(maxlen):
                        colname = " ".join([header_parts[r][i] for r in range(len(header_parts)) if header_parts[r][i]])
                        colname = re.sub(r"\s+", " ", colname).strip()
                        columns.append(colname if colname else f"COL_{i+1}")

                    # UNIKKAN KOLOM!
                    columns = make_columns_unique(columns)
                    n_cols = len(columns)

                    # --- Parsing data
                    parsed = []
                    for ln in lines[header_idx+len(header_lines):]:
                        if "SUB TOTAL" in ln.upper():
                            continue
                        row = ln.replace("Â", "").replace("Ã", "").replace("°", "")
                        if row.startswith(delimiter):
                            row = row[len(delimiter):]
                        parts = [p.strip().replace("³","") for p in row.split(delimiter)]
                        if len(parts)<n_cols:
                            parts += [""]*(n_cols-len(parts))
                        elif len(parts)>n_cols:
                            parts = parts[:n_cols]
                        # Data row valid: minimal 2 kolom pertama tidak kosong
                        if not all([x == "" for x in parts[:2]]):
                            parsed.append(parts)

                    if not parsed:
                        df = pd.DataFrame(columns=columns)
                    else:
                        df = pd.DataFrame(parsed, columns=columns)
                        nik_col = [c for c in df.columns if "NIK" in c.upper()]
                        nama_col = [c for c in df.columns if "NAMA" in c.upper()]
                        if nik_col and nama_col:
                            nik_col = nik_col[0]
                            nama_col = nama_col[0]
                            nik_asli = df[nik_col].astype(str)
                            df[nik_col] = nik_asli.str.extract(r"^([A-Za-z0-9]+)")
                            if "Kode" not in df.columns:
                                df['Kode'] = nik_asli.str.extract(r"^[A-Za-z0-9]+\s+(\d+)")
                                df['Kode'] = df['Kode'].fillna("")
                            df[nama_col] = df[nama_col].astype(str).str.strip()
                            df = df[
                                df[nik_col].notna() &
                                (df[nik_col].astype(str).str.strip() != "") &
                                (df[nik_col].astype(str).str.upper() != "NIK") &
                                (df[nama_col].astype(str).str.strip() != "") &
                                (df[nama_col].astype(str).str.upper() != "NAMA") &
                                (df.drop([nik_col, nama_col], axis=1).apply(lambda row: any([str(x).strip() != "" for x in row]), axis=1))
                            ].copy()
                            df.reset_index(drop=True, inplace=True)
                            cols = df.columns.tolist()
                            if 'Kode' in cols and nik_col in cols:
                                nik_idx = cols.index(nik_col)
                                cols.remove('Kode')
                                cols.insert(nik_idx+1, 'Kode')
                                df = df[cols]
                        df = df[[c for c in df.columns if not c.startswith("COL_")]]
                        # Hapus kolom 'c' jika ada
                        if 'c' in df.columns:
                            df = df.drop(columns=['c'])
                        # Hapus kolom 'Kode' jika ini absensi
                        is_absensi = (
                            (split_header and "absensi" in split_header.lower())
                            or any("absensi" in str(h).lower() for h in columns)
                            or any("absensi" in str(section).lower() for section in sections)
                        )
                        if is_absensi and "Kode" in df.columns:
                            df = df.drop(columns=["Kode"])
                        # --- Konversi kolom angka
                        def is_number_col(series):
                            cleaned = series.astype(str)\
                                .str.replace(",", "", regex=False)\
                                .str.replace(" ", "", regex=False)\
                                .replace("", np.nan)
                            cleaned = cleaned[~cleaned.isna()]
                            def isfloat(x):
                                try:
                                    float(x)
                                    return True
                                except Exception:
                                    return False
                            return len(cleaned) > 0 and all(isfloat(x) for x in cleaned)
                        for col in df.columns:
                            if col.strip().lower() in ["nik", "kode", "nama"]:
                                continue
                            if is_number_col(df[col]):
                                df[col] = pd.to_numeric(
                                    df[col].astype(str)
                                    .str.replace(",", "", regex=False)
                                    .str.replace(" ", "", regex=False),
                                    errors="coerce"
                                )
                    # END parsing

                sheet_name = f"Page_{idx}" if len(f"Page_{idx}")<=31 else f"Pg_{idx}"
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                preview_dfs.append((sheet_name, df))

                # Format kolom di Excel: numerik dengan #,##0
                wb = writer.book
                ws = writer.sheets[sheet_name]
                fmt_text = wb.add_format({"num_format":"@"})
                for i, c in enumerate(df.columns):
                    if c.strip().lower() in ["nik", "kode", "nama"]:
                        ws.set_column(i,i,16,fmt_text)
                    else:
                        ws.set_column(i,i,16, wb.add_format({'num_format': '#,##0'}))

    # Download utama Excel dan ZIP (zip di kiri, excel kanan)
    col1, col2 = st.columns(2)
    with col1:
        with open(zip_path,"rb") as fzip:
            st.download_button(
                "📦 Download Semua File TXT (.zip)",
                fzip,
                file_name="hasil_split_gaji_karyawan_txt.zip",
                mime="application/zip"
            )
    with col2:
        with open(excel_path,"rb") as fexcel:
            st.download_button(
                "📥 Download File Excel (.xlsx)",
                fexcel,
                file_name="hasil_split_gaji_karyawan.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    # Preview Sheet
    for sheet_name, df in preview_dfs:
        with st.expander(f"🔍 Preview Sheet '{sheet_name}'"):
            if df.empty:
                st.write("_Sheet kosong_")
            else:
                df_preview = get_preview_df(df)
                st.dataframe(df_preview, use_container_width=True)
                # Download per sheet
                single_excel_name = st.text_input(
                    f"Nama file untuk sheet '{sheet_name}' (.xlsx)", value=f"{sheet_name}.xlsx", key=f"fname_{sheet_name}"
                )
                single_excel_path = os.path.join(tmpdirname, f"{sheet_name}.xlsx")
                with pd.ExcelWriter(single_excel_path, engine="xlsxwriter") as single_writer:
                    df.to_excel(single_writer, index=False, sheet_name=sheet_name)
                    wb = single_writer.book
                    ws = single_writer.sheets[sheet_name]
                    fmt_text = wb.add_format({"num_format": "@"})
                    for i, c in enumerate(df.columns):
                        if c.strip().lower() in ["nik", "kode", "nama"]:
                            ws.set_column(i,i,16,fmt_text)
                        else:
                            ws.set_column(i,i,16, wb.add_format({'num_format': '#,##0'}))
                with open(single_excel_path, "rb") as single_excel_file:
                    st.download_button(
                        label=f"⬇️ Download Sheet '{sheet_name}' (.xlsx)",
                        data=single_excel_file,
                        file_name=single_excel_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

    # Preview isi file txt mentah
    for fn_txt, fp_txt, content in filepaths:
        with st.expander(f"📄 Preview {fn_txt}"):
            preview = "\n".join(content.splitlines()[:20])
            st.text(preview)
            with open(fp_txt, "rb") as ftxt:
                st.download_button(
                    label=f"⬇️ Download {fn_txt}",
                    data=ftxt,
                    file_name=fn_txt,
                    mime="text/plain"
                )
