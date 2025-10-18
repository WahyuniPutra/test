import io
import re
from pathlib import Path

import pandas as pd
import streamlit as st

# ================== SET TEMPLATE PATH LOKAL ==================
TEMPLATE_PATH = Path(r"C:\Users\UsEr\Downloads\tess\template-import-participants.xlsx")
# =============================================================

st.set_page_config(page_title="Import Participants Filler", page_icon="🧑‍🎓", layout="wide")
st.title("🧑‍🎓 Import Participants Filler")
st.caption("Pindahkan data dari Excel sumber ke template import peserta — aman untuk NISN dengan nol di depan.")

# ================== UTIL FUNCS ==================
def read_first_sheet_stream(file, force_text: bool = False) -> pd.DataFrame:
    """Baca worksheet pertama dari file uploader Streamlit."""
    kwargs = {}
    if force_text:
        kwargs.update(dict(dtype=str, keep_default_na=False))
    data = file.read()
    bio = io.BytesIO(data)
    return pd.read_excel(bio, sheet_name=0, engine="openpyxl", **kwargs)

def read_first_sheet_path(path: Path, force_text: bool = False) -> pd.DataFrame:
    """Baca worksheet pertama dari path lokal."""
    kwargs = {}
    if force_text:
        kwargs.update(dict(dtype=str, keep_default_na=False))
    return pd.read_excel(path, sheet_name=0, engine="openpyxl", **kwargs)

def find_col(df: pd.DataFrame, patterns):
    for col in df.columns:
        for pat in patterns:
            if re.search(pat, str(col).strip(), flags=re.I):
                return col
    return None

def ensure_len(df: pd.DataFrame, n: int):
    if len(df) < n:
        extra = pd.DataFrame({c: [None]*(n-len(df)) for c in df.columns})
        return pd.concat([df, extra], ignore_index=True)
    return df

def set_col(df, col_guess, fallback, values):
    col = col_guess or fallback
    if col not in df.columns:
        df[col] = None
    df[col] = (list(values) + [None]*max(0, len(df)-len(values)))[:len(df)]
    return col

def derive_school_from_filename(name: str) -> str:
    stem = Path(name).stem
    s = stem.lower()
    s = re.sub(r"\([^)]*\)", " ", s)
    s = re.sub(r"[_\-\.]+", " ", s)
    stopwords = {
        "download","biodata","template","import","participants",
        "participant","peserta","particioants","data","final","rev",
        "siswa","kelas","ruangan","nisn","copy"
    }
    tokens = [t for t in s.split() if t and t not in stopwords and not t.isdigit()]
    if not tokens:
        return ""
    guess = " ".join(tokens).strip()
    return " ".join(w.capitalize() for w in guess.split())

def get_values_or_constant(constant_value: str, src_df: pd.DataFrame, patterns_for_src_col, n_rows: int):
    if constant_value:
        return [constant_value] * n_rows
    src_col = find_col(src_df, patterns_for_src_col)
    if src_col:
        return src_df[src_col].fillna("").astype(str).tolist()
    return [None] * n_rows

def build_output(src_df: pd.DataFrame, tmpl_df: pd.DataFrame,
                 class_code: str, room_code: str, session_code: str, school_name: str):
    # Deteksi kolom sumber penting
    nisn_col = find_col(src_df, [r"\bnisn\b"])
    name_col = find_col(src_df, [r"\bnama\b", r"\bnama\s*lengkap\b", r"\bnama\s*peserta\b", r"\bname\b"])
    if nisn_col is None:
        raise ValueError(f"Kolom NISN tidak ditemukan di sumber. Kolom tersedia: {list(src_df.columns)}")
    if name_col is None:
        raise ValueError(f"Kolom Nama tidak ditemukan di sumber. Kolom tersedia: {list(src_df.columns)}")

    # Ambil apa adanya (leading zero aman karena src_df sudah dtype=str)
    nisn_vals = src_df[nisn_col].fillna("").astype(str).str.strip().tolist()
    name_vals = src_df[name_col].fillna("").astype(str).str.strip().tolist()

    # Samakan panjang template = jumlah peserta
    tmpl_df = ensure_len(tmpl_df, len(nisn_vals))

    # Deteksi kolom di template
    id_col        = find_col(tmpl_df, [r"\bid\b", r"\bid\s*peserta\b", r"\bid[_\s-]*peserta\b"])
    username_col  = find_col(tmpl_df, [r"\busername\b"])
    password_col  = find_col(tmpl_df, [r"\bpassword\b"])
    name_out_col  = find_col(tmpl_df, [r"\bnama\b", r"\bname\b", r"\bnama\s*lengkap\b"])
    kelas_col     = find_col(tmpl_df, [r"\bkode\s*kelas\b", r"\bkelas\b"])
    ruangan_col   = find_col(tmpl_df, [r"\bkode\s*ruangan\b", r"\bruangan\b"])
    sekolah_col   = find_col(tmpl_df, [r"\bsekolah\b"])
    sesi_col      = find_col(tmpl_df, [r"\bsesi\b"])

    # Isi kolom utama
    set_col(tmpl_df, id_col,       "ID PESERTA", nisn_vals)
    set_col(tmpl_df, username_col, "USERNAME",   nisn_vals)
    set_col(tmpl_df, password_col, "PASSWORD",   nisn_vals)
    set_col(tmpl_df, name_out_col, "NAMA",       name_vals)

    # KODE KELAS, KODE RUANGAN
    kelas_vals   = get_values_or_constant(class_code, src_df, [r"\bkelas\b", r"\bkode\s*kelas\b"], len(nisn_vals))
    ruangan_vals = get_values_or_constant(room_code,  src_df, [r"\bruang(an)?\b", r"\bkode\s*ruangan\b"], len(nisn_vals))
    set_col(tmpl_df, kelas_col,   "KODE KELAS",   kelas_vals)
    set_col(tmpl_df, ruangan_col, "KODE RUANGAN", ruangan_vals)

    # SESI
    sesi_vals = get_values_or_constant(session_code, src_df, [r"\bsesi\b"], len(nisn_vals))
    set_col(tmpl_df, sesi_col, "SESI", sesi_vals)

    # SEKOLAH
    set_col(tmpl_df, sekolah_col, "SEKOLAH", [school_name] * len(nisn_vals))

    return tmpl_df

def to_excel_bytes(df: pd.DataFrame, sheet_name: str = "import") -> bytes:
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    bio.seek(0)
    return bio.getvalue()

# ================== UI ==================
st.sidebar.header("⚙️ Pengaturan")
uploaded_src = st.sidebar.file_uploader("1) Upload Excel Sumber (Download_Biodata…)", type=["xlsx"])

# Tampilkan status template path
with st.sidebar.expander("📄 Template yang dipakai (fixed)"):
    st.code(str(TEMPLATE_PATH))
    if not TEMPLATE_PATH.exists():
        st.error("Template tidak ditemukan di path tersebut. Pastikan file ada & path benar.")

# Parameter konstan
class_code = st.sidebar.text_input("KODE KELAS (opsional, biarkan kosong untuk ambil dari sumber)", value="")
room_code  = st.sidebar.text_input("KODE RUANGAN (opsional, biarkan kosong untuk ambil dari sumber)", value="")
session_code = st.sidebar.text_input("SESI (opsional, biarkan kosong untuk ambil dari sumber)", value="")

# Sekolah: default ditebak dari nama file sumber
default_school = ""
if uploaded_src is not None:
    default_school = derive_school_from_filename(uploaded_src.name) or ""
school_name = st.sidebar.text_input("SEKOLAH", value=default_school)

st.markdown("#### 1) Upload **Excel sumber** di sidebar. Template otomatis dibaca dari path lokal di atas, lalu klik **Generate**.")

generate = st.button("🚀 Generate Hasil Import")

if generate:
    if not uploaded_src:
        st.error("Mohon upload **Excel Sumber** terlebih dahulu.")
    elif not TEMPLATE_PATH.exists():
        st.error("Template tidak ditemukan. Periksa path di sidebar.")
    else:
        try:
            # Baca sumber (paksa string) & template (normal) dari PATH LOKAL
            src_df = read_first_sheet_stream(uploaded_src, force_text=True)
            tmpl_df = read_first_sheet_path(TEMPLATE_PATH, force_text=False)

            # Build output
            out_df = build_output(
                src_df=src_df,
                tmpl_df=tmpl_df.copy(),
                class_code=class_code,
                room_code=room_code,
                session_code=session_code,
                school_name=school_name,
            )

            # Preview
            st.success("Berhasil menyusun data import.")
            c1, c2 = st.columns(2)
            with c1:
                st.subheader("👈 Preview Sumber (top 20)")
                st.dataframe(src_df.head(20), use_container_width=True)
            with c2:
                st.subheader("👉 Preview Hasil Import (top 20)")
                st.dataframe(out_df.head(20), use_container_width=True)

            # Unduh
            xlsx_bytes = to_excel_bytes(out_df, sheet_name="import")
            suggested_name = f"{school_name or 'hasil'}_filled.xlsx"
            st.download_button(
                label="💾 Download Excel Hasil",
                data=xlsx_bytes,
                file_name=suggested_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

            with st.expander("ℹ️ Catatan & Deteksi Otomatis"):
                st.markdown("""
- **NISN** → **USERNAME**, **PASSWORD**, **ID PESERTA** (leading zero aman).
- **NAMA** → dari kolom `Nama / Nama Lengkap / Nama Peserta / Name`.
- **KODE KELAS**, **KODE RUANGAN**, **SESI**:
  - Jika input konstanta di sidebar diisi → pakai konstanta untuk semua baris.
  - Jika dikosongkan, dan kolomnya ada di sumber → pakai data sumber.
  - Jika tidak ada keduanya → dibuat kolomnya tapi nilainya kosong.
- **SEKOLAH** → dari field di sidebar (default ditebak dari nama file sumber).
                """)
        except Exception as e:
            st.error(f"Gagal memproses: {e}")

# Footer
st.markdown("---")
st.caption("Made with ❤️ Streamlit • Pandas • openpyxl — menjaga NISN dengan leading zero tetap utuh.")
