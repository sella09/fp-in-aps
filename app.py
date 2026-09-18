import streamlit as st
import pdfplumber
import re
import pandas as pd
from io import BytesIO
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Ekstraktor Faktur Pajak", layout="wide")
st.title("📄 Ekstraktor Faktur Pajak Indonesia")
st.caption("Upload PDF faktur → Excel berisi Rekap + 1 sheet per faktur, nominal sebagai angka asli.")


# =========================================================
# HELPER
# =========================================================
def parse_angka(s):
    if not s:
        return None
    s = str(s).strip().replace("Rp", "").replace(" ", "")
    s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def rupiah(x):
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return "-"
    return f"Rp {x:,.0f}".replace(",", ".")


def safe_sheet_name(name, used):
    base = re.sub(r"[\\/*?:\[\]]", "_", str(name))[:28]
    sheet = base or "Sheet"
    i = 1
    while sheet in used:
        suffix = f"_{i}"
        sheet = base[:31 - len(suffix)] + suffix
        i += 1
    used.add(sheet)
    return sheet


# =========================================================
# EKSTRAKSI
# =========================================================
def ekstrak_header(teks):
    data = {}
    m = re.search(r"Nomor Seri Faktur Pajak:\s*(\d+)", teks)
    data["nomor_seri"] = m.group(1) if m else None

    npwp_all = re.findall(r"NPWP\s*:\s*(\d{15,16})", teks)
    data["npwp_penjual"] = npwp_all[0] if len(npwp_all) > 0 else None
    data["npwp_pembeli"] = npwp_all[1] if len(npwp_all) > 1 else None

    m = re.search(r"Pengusaha Kena Pajak:\s*\n?\s*Nama\s*:\s*(.+)", teks)
    data["nama_penjual"] = m.group(1).strip() if m else None

    m = re.search(r"Pembeli Barang Kena Pajak/Penerima Jasa Kena Pajak:\s*\n?\s*Nama\s*:\s*(.+)", teks)
    data["nama_pembeli"] = m.group(1).strip() if m else None

    m = re.search(r"Pengusaha Kena Pajak:.*?Alamat\s*:\s*(.+)", teks, re.DOTALL)
    data["alamat_penjual"] = m.group(1).strip().split("\n")[0] if m else None

    m = re.search(r"Pembeli Barang Kena Pajak/Penerima Jasa Kena Pajak:.*?Alamat\s*:\s*(.+)", teks, re.DOTALL)
    data["alamat_pembeli"] = m.group(1).strip().split("\n")[0] if m else None

    m = re.search(r"Dasar Pengenaan Pajak\s*([\d.,]+)", teks)
    data["dpp"] = parse_angka(m.group(1)) if m else None

    m = re.search(r"Jumlah PPN[^0-9]*([\d.,]+)", teks)
    data["ppn"] = parse_angka(m.group(1)) if m else None

    m = re.search(r"Jumlah PPnBM[^0-9]*([\d.,]+)", teks)
    data["ppnbm"] = parse_angka(m.group(1)) if m else None

    m = re.search(r"Harga Jual / Penggantian / Uang Muka / Termin\s*([\d.,]+)", teks)
    data["harga_jual_total"] = parse_angka(m.group(1)) if m else None

    m = re.search(r"JAKARTA[^,]*,\s*(\d{1,2}\s+\w+\s+\d{4})", teks)
    data["tanggal"] = m.group(1) if m else None

    return data


def ekstrak_barang(teks):
    items = []
    pola = re.compile(
        r"(?P<nama>[A-Za-z0-9\s\-/]+?)\s*-\s*Part No\s*:\s*(?P<part>\S+)\s*"
        r"Rp\s*(?P<harga>[\d.,]+)\s*x\s*(?P<qty>[\d.,]+)\s*(?P<satuan>\w+).*?"
        r"(?P<subtotal>[\d.,]+)\s*$",
        re.MULTILINE
    )
    for m in pola.finditer(teks):
        items.append({
            "nama_barang": m.group("nama").strip(),
            "part_no": m.group("part").strip(),
            "harga_satuan": parse_angka(m.group("harga")),
            "qty": parse_angka(m.group("qty")),
            "satuan": m.group("satuan").strip(),
            "subtotal": parse_angka(m.group("subtotal")),
        })
    return items


# =========================================================
# PROSES PDF
# =========================================================
def proses_pdf(file):
    with pdfplumber.open(file) as pdf:
        teks = "\n".join(page.extract_text() or "" for page in pdf.pages)
    header = ekstrak_header(teks)
    header["nama_file"] = file.name
    barang = ekstrak_barang(teks)
    return header, barang, teks


# =========================================================
# UI
# =========================================================
uploaded_files = st.file_uploader(
    "Upload file PDF faktur pajak (bisa banyak)",
    type=["pdf"],
    accept_multiple_files=True,
)

if uploaded_files:
    if st.button("🚀 Proses & Validasi", type="primary"):
        hasil = []
        progress = st.progress(0)

        for i, f in enumerate(uploaded_files):
            try:
                h, b, teks = proses_pdf(f)
                hasil.append({"header": h, "barang": b, "teks": teks, "nama": f.name})
            except Exception as e:
                st.warning(f"Gagal memproses {f.name}: {e}")
            progress.progress((i + 1) / len(uploaded_files))

        st.success("Selesai!")

        # ---- DEBUG: tampilkan teks mentah ----
        st.header("🔍 Debug: Teks dari PDF")
        st.info("Copy atau screenshot isi teks di bawah ini, kirim ke developer untuk penyesuaian regex.")
        for r in hasil:
            with st.expander(f"📄 {r['nama']} — {len(r['teks'])} karakter"):
                st.text(r["teks"])

        # ---- Ringkasan cepat ----
        st.header("📊 Ringkasan Ekstraksi")
        for r in hasil:
            h = r["header"]
            st.markdown(f"**{r['nama']}**")
            st.json(h)
            if r["barang"]:
                st.dataframe(pd.DataFrame(r["barang"]))
            else:
                st.warning("⚠️ Tidak ada barang yang terekstrak dari PDF ini.")

else:
    st.info("Silakan upload file PDF faktur pajak terlebih dahulu.")
