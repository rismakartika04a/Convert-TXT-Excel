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
    ["", "GAJI KARYAWAN TETAP", "GAJI KARYAWAN KONTRAK"]
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

with tempfile.TemporaryDirectory() as tmpdirname:
    zip_path = os.path.join(tmpdirname, "hasil_split_gaji_karyawan_txt.zip")
    zipf = zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED)
    excel_path = os.path.join(tmpdirname, "hasil_split_gaji_karyawan.xlsx")
    writer = pd.ExcelWriter(excel_path, engine="xlsxwriter")
    preview_dfs = []
    filepaths = []

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
            header_idx = None
            for i, ln in enumerate(lines):
                if "NIK" in ln.upper() and "NAMA" in ln.upper():
                    header_idx = i
                    break
            if header_idx is None:
                header_idx = 0

            header1 = lines[header_idx].replace("Â", "").replace("Ã", "").replace("°", "")
            header2 = ""
            if header_idx+1 < len(lines):
                header2 = lines[header_idx+1].replace("Â", "").replace("Ã", "").replace("°", "")

            parts1 = [p.strip() for p in header1.split(delimiter)]
            parts2 = [p.strip() for p in header2.split(delimiter)] if header2 else [""]*len(parts1)
            maxlen = max(len(parts1), len(parts2))
            while len(parts1)<maxlen: parts1.append("")
            while len(parts2)<maxlen: parts2.append("")

            # Gabung header dua baris, lalu buat unik
            raw_cols = []
            for h1, h2 in zip(parts1, parts2):
                if h1 and h2: nm = f"{h1} {h2}".strip()
                else: nm = h1 or h2
                nm = re.sub(r"\s+", " ", nm).strip()
                raw_cols.append(nm if nm else "COL")
            seen = {}
            columns = []
            for col in raw_cols:
                if col not in seen:
                    seen[col]=1
                    columns.append(col)
                else:
                    seen[col] +=1
                    columns.append(f"{col}_{seen[col]}")
            n_cols = len(columns)

            # --- Parsing data
            parsed = []
            for ln in lines[header_idx+2:]:
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
                parsed.append(parts)

            if not parsed:
                df = pd.DataFrame(columns=columns)
            else:
                df = pd.DataFrame(parsed, columns=columns)
                nik_col = [c for c in df.columns if c.strip().upper() == 'NIK']
                nama_col = [c for c in df.columns if c.strip().upper() == 'NAMA']
                if nik_col and nama_col:
                    nik_col = nik_col[0]
                    nama_col = nama_col[0]
                    nik_asli = df[nik_col].astype(str)
                    df[nik_col] = nik_asli.str.extract(r"^([A-Za-z0-9]+)")
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
                    # --- Susun ulang kolom agar Kode setelah NIK ---
                    cols = df.columns.tolist()
                    if 'Kode' in cols and nik_col in cols:
                        nik_idx = cols.index(nik_col)
                        cols.remove('Kode')
                        cols.insert(nik_idx+1, 'Kode')
                        df = df[cols]
                # Hapus kolom COL, COL_2 dst jika ada
                df = df[[c for c in df.columns if not c.startswith("COL")]]

                # PATCH: Convert kolom angka
                def is_number_col(series):
                    cleaned = series.astype(str).str.replace(",", "").str.replace(" ", "").str.replace("-", "").replace("", np.nan)
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
                        df[col] = pd.to_numeric(df[col].astype(str).str.replace(",", "").str.replace(" ", "").str.replace("-", ""), errors="coerce")

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

    writer.close()
    zipf.close()

    # Download utama Excel dan ZIP
    col1, col2 = st.columns(2)
    with col1:
        with open(excel_path,"rb") as fexcel:
            st.download_button(
                "📥 Download File Excel (.xlsx)",
                fexcel,
                file_name="hasil_split_gaji_karyawan.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    with col2:
        with open(zip_path,"rb") as fzip:
            st.download_button(
                "📦 Download Semua File TXT (.zip)",
                fzip,
                file_name="hasil_split_gaji_karyawan_txt.zip",
                mime="application/zip"
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
