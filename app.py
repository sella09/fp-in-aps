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
# EKSTRAKSI HEADER
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

    # DPP - berbagai varian penulisan
    m = re.search(r"Dasar Pengenaan Pajak\s*([\d.,]+)", teks)
    data["dpp"] = parse_angka(m.group(1)) if m else None

    # PPN - fleksibel terhadap spasi dan newline
    m = re.search(r"Jumlah PPN[^0-9]*([\d.,]+)", teks, re.DOTALL)
    data["ppn"] = parse_angka(m.group(1)) if m else None

    # PPnBM - 3 varian: PPnBM, PPNbM, PnPbM
    m = re.search(r"Jumlah\s+P+[nN]?[bB]?[mM]?[^0-9]*([\d.,]+)", teks)
    data["ppnbm"] = parse_angka(m.group(1)) if m else None

    m = re.search(r"Harga Jual / Penggantian / Uang Muka / Termin\s*([\d.,]+)", teks)
    data["harga_jual_total"] = parse_angka(m.group(1)) if m else None

    # Tanggal - fleksibel kota apapun
    m = re.search(r"(?:JAKARTA|TANGERANG|KAB\.\s*TANGERANG|BOGOR|DEPOK|BEKASI|SURABAYA|BANDUNG)[^,]*,\s*(\d{1,2}\s+\w+\s+\d{4})", teks)
    data["tanggal"] = m.group(1) if m else None

    return data


# =========================================================
# EKSTRAKSI BARANG (VERSI FIX MULTI-FORMAT)
# =========================================================
def ekstrak_barang(teks):
    """
    Deteksi 2 format:
    A) Satu baris panjang: "1 620000 Seragam Kerja Rp 230.000,00 x 5,00 Set ... 1.150.000,00"
    B) Multi-baris (PDF terpisah):
       "1 000000 Washer - Part No : 11210753"
       "Rp 23.441,25 x 10,00 Unit"
       "Potongan Harga = Rp 0,00"
       "PPnBM (0,00%) = Rp 0,00"
       "234.412,50"
    """
    items = []
    posisi_part = [(m.start(), m) for m in re.finditer(r"-\s*Part No\s*:\s*(\S+)", teks)]

    # Apakah ada "- Part No" (format SDLG-style)?
    if posisi_part:
        m_end = re.search(r"Harga Jual / Penggantian", teks)
        batas_akhir = m_end.start() if m_end else len(teks)

        for i, (start, m) in enumerate(posisi_part):
            end = posisi_part[i + 1][0] if i + 1 < len(posisi_part) else batas_akhir
            blok = teks[start:end]
            part_no = m.group(1).strip()

            # Nama barang: ambil kata sebelum "- Part No"
            before = teks[max(0, start - 100):start]
            m_nama = re.search(r"(?:^|\s)\d{1,2}\s+\d{6}\s+([A-Za-z][A-Za-z0-9\s\-/\.]*?)\s*$", before)
            if not m_nama:
                m_nama = re.search(r"(\S+(?:\s+\S+)*?)\s*$", before)
            nama = m_nama.group(1).strip() if m_nama else ""

            m2 = re.search(r"Rp\s*([\d.,]+)\s*x\s*([\d.,]+)\s*(\w+)", blok)
            harga = parse_angka(m2.group(1)) if m2 else None
            qty = parse_angka(m2.group(2)) if m2 else None
            satuan = m2.group(3).strip() if m2 else None

            # Subtotal: angka setelah varian PPnBM
            m3 = re.search(
                r"[Pp]+[Nn]?[Bb]?[Mm]?\s*\(\s*0[,.]00%\s*\)\s*=\s*Rp\s*[\d.,]+\s*([\d.,]+)",
                blok
            )
            if m3:
                subtotal = parse_angka(m3.group(1))
            else:
                angka = re.findall(r"(\d{1,3}(?:\.\d{3})+(?:,\d{2})?)", blok)
                subtotal = parse_angka(angka[-1]) if angka else None

            items.append({
                "nama_barang": nama,
                "part_no": part_no,
                "harga_satuan": harga,
                "qty": qty,
                "satuan": satuan,
                "subtotal": subtotal,
            })
        return items

    # Fallback: format NADE (tanpa Part No) - pola satu baris panjang
    # Cari: No Kode Nama Rp harga x qty Satuan ... subtotal
    pola = re.compile(
        r"(?P<no>\d{1,3})\s+(?P<kode>\d{6})\s+"
        r"(?P<nama>[A-Za-z][A-Za-z0-9\s\-/\.]*?)\s+"
        r"Rp\s*(?P<harga>[\d.,]+)\s*x\s*(?P<qty>[\d.,]+)\s*(?P<satuan>\w+).*?"
        r"(?P<subtotal>\d{1,3}(?:\.\d{3})+(?:,\d{2})?)",
        re.DOTALL
    )
    for m in pola.finditer(teks):
        items.append({
            "nama_barang": m.group("nama").strip(),
            "part_no": m.group("kode").strip(),
            "harga_satuan": parse_angka(m.group("harga")),
            "qty": parse_angka(m.group("qty")),
            "satuan": m.group("satuan").strip(),
            "subtotal": parse_angka(m.group("subtotal")),
        })
    return items


# =========================================================
# VALIDASI (VERSI LENTUR)
# =========================================================
TOLERANSI = 1.0


def validasi_satu_faktur(header, barang):
    rows = []
    dpp = header.get("dpp")
    ppn = header.get("ppn")
    ppnbm = header.get("ppnbm") or 0
    harga_jual = header.get("harga_jual_total")

    total_item = 0
    for i, b in enumerate(barang, 1):
        hs, q, st_ = b.get("harga_satuan"), b.get("qty"), b.get("subtotal")
        if hs is None or q is None or st_ is None:
            status, ket, selisih = "❌", "Data tidak lengkap", None
        else:
            hitung = round(hs * q, 2)
            selisih = round(st_ - hitung, 2)
            status = "✅" if abs(selisih) <= TOLERANSI else "❌"
            ket = "OK" if status == "✅" else f"Selisih {rupiah(selisih)}"
            total_item += st_
        rows.append({
            "Cek": f"Item {i}: {b.get('nama_barang')}",
            "Nilai": f"{rupiah(hs)} × {q}",
            "Selisih": selisih,
            "Status": status,
            "Keterangan": ket,
        })

    # Total item vs Harga Jual
    if harga_jual is not None and total_item > 0:
        selisih = round(harga_jual - total_item, 2)
        status = "✅" if abs(selisih) <= TOLERANSI else "❌"
        ket = "OK" if status == "✅" else f"Selisih {rupiah(selisih)}"
    else:
        selisih, status, ket = None, "⚠️", "Tidak bisa dicek"
    rows.append({
        "Cek": "Total item vs Harga Jual",
        "Nilai": f"{rupiah(total_item)} vs {rupiah(harga_jual)}",
        "Selisih": selisih,
        "Status": status, "Keterangan": ket,
    })

    # DPP + PPN + PPnBM vs Harga Jual (hanya cek kalau DPP mendekati Harga Jual)
    # (kalau DPP Nilai Lain, tidak dicek - hanya info)
    if dpp is not None and harga_jual is not None and harga_jual > 0:
        rasio = dpp / harga_jual
        if 0.95 <= rasio <= 1.05:
            # DPP normal
            if ppn is not None:
                hitung = round(dpp + ppn + ppnbm, 2)
                selisih = round(harga_jual - hitung, 2)
                status = "✅" if abs(selisih) <= TOLERANSI else "❌"
                ket = "OK" if status == "✅" else f"Selisih {rupiah(selisih)}"
            else:
                selisih, status, ket = None, "⚠️", "PPN tidak terbaca"
        else:
            # DPP Nilai Lain - info saja
            selisih, status = None, "ℹ️"
            ket = f"DPP Nilai Lain (rasio {rasio:.2%}) - validasi dilewati"
    else:
        selisih, status, ket = None, "⚠️", "Tidak bisa dicek"
    rows.append({
        "Cek": "DPP + PPN + PPnBM vs Harga Jual",
        "Nilai": f"{rupiah(dpp)} + {rupiah(ppn)} + {rupiah(ppnbm)} vs {rupiah(harga_jual)}",
        "Selisih": selisih,
        "Status": status, "Keterangan": ket,
    })

    # Tarif PPN efektif (hanya info, bukan validasi)
    if dpp and ppn is not None and dpp > 0:
        tarif = ppn / dpp
        status = "ℹ️"
        ket = f"Tarif efektif {tarif:.2%}"
    else:
        status, ket = "⚠️", "Tidak bisa dicek"
    rows.append({
        "Cek": "Tarif PPN Efektif",
        "Nilai": f"PPN {rupiah(ppn)} ÷ DPP {rupiah(dpp)}",
        "Selisih": None,
        "Status": status, "Keterangan": ket,
    })

    # Kelengkapan data wajib
    wajib = ["nomor_seri", "npwp_penjual", "npwp_pembeli", "tanggal", "dpp", "ppn"]
    kosong = [k for k in wajib if not header.get(k)]
    status = "✅" if not kosong else "❌"
    ket = "OK" if not kosong else f"Kosong: {', '.join(kosong)}"
    rows.append({
        "Cek": "Kelengkapan Data Wajib",
        "Nilai": ", ".join(wajib),
        "Selisih": None,
        "Status": status, "Keterangan": ket,
    })

    # Status akhir: hanya ❌ kalau ada ❌, kalau hanya ℹ️/⚠️ → tetap ✅
    if any(r["Status"] == "❌" for r in rows):
        akhir = "❌ Gagal"
    elif any(r["Status"] == "⚠️" for r in rows):
        akhir = "⚠️ Perlu dicek"
    else:
        akhir = "✅ Valid"

    rows.append({"Cek": "STATUS AKHIR", "Nilai": "", "Selisih": None,
                 "Status": akhir, "Keterangan": ""})
    return rows


# =========================================================
# PROSES PDF
# =========================================================
def proses_pdf(file):
    with pdfplumber.open(file) as pdf:
        teks = "\n".join(page.extract_text() or "" for page in pdf.pages)
    header = ekstrak_header(teks)
    header["nama_file"] = file.name
    barang = ekstrak_barang(teks)
    return header, barang


# =========================================================
# EXCEL HELPER
# =========================================================
def tulis_df(writer, df, sheet, startrow=0, startcol=0,
             kolom_angka=None, header_fill="D9E1F2"):
    df.to_excel(writer, sheet_name=sheet, index=False,
                startrow=startrow, startcol=startcol)
    ws = writer.sheets[sheet]
    kolom_angka = kolom_angka or []

    for col_name in kolom_angka:
        if col_name not in df.columns:
            continue
        idx = df.columns.get_loc(col_name) + startcol + 1
        letter = get_column_letter(idx)
        for row in range(startrow + 2, startrow + 2 + len(df)):
            ws[f"{letter}{row}"].number_format = "#,##0"

    for col in range(startcol + 1, startcol + 1 + len(df.columns)):
        c = ws.cell(row=startrow + 1, column=col)
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor=header_fill)
        c.alignment = Alignment(horizontal="center", vertical="center")

    return startrow + len(df) + 1


def tulis_judul(writer, sheet, text, row):
    ws = writer.sheets[sheet]
    c = ws.cell(row=row + 1, column=1, value=text)
    c.font = Font(bold=True, size=12)
    return row + 2


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
        status = st.empty()

        for i, f in enumerate(uploaded_files):
            status.text(f"Memproses: {f.name} ({i+1}/{len(uploaded_files)})")
            try:
                h, b = proses_pdf(f)
                v = validasi_satu_faktur(h, b)
                hasil.append({"header": h, "barang": b, "validasi": v})
            except Exception as e:
                st.warning(f"Gagal memproses {f.name}: {e}")
            progress.progress((i + 1) / len(uploaded_files))
        status.text("✅ Selesai!")

        st.session_state["hasil"] = hasil

        total = len(hasil)
        valid = sum(1 for r in hasil if r["validasi"][-1]["Status"] == "✅ Valid")
        perlu = sum(1 for r in hasil if r["validasi"][-1]["Status"] == "⚠️ Perlu dicek")
        gagal = sum(1 for r in hasil if r["validasi"][-1]["Status"] == "❌ Gagal")

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total Faktur", total)
        c2.metric("✅ Valid", valid)
        c3.metric("⚠️ Perlu Dicek", perlu)
        c4.metric("❌ Gagal", gagal)

        for r in hasil:
            h = r["header"]
            st.markdown("---")
            st.subheader(f"📄 {h.get('nomor_seri') or h.get('nama_file')}")
            st.caption(f"Pembeli: {h.get('nama_pembeli')} | Tanggal: {h.get('tanggal')}")
            st.markdown("**Detail Barang**")
            st.dataframe(pd.DataFrame(r["barang"]), use_container_width=True)
            st.markdown("**Validasi**")
            st.dataframe(pd.DataFrame(r["validasi"]), use_container_width=True)

        # ================= EXPORT EXCEL =================
        buffer = BytesIO()
        used_names = set()

        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            rekap_sheet = safe_sheet_name("Rekap", used_names)
            pd.DataFrame([["REKAP FAKTUR PAJAK"]]).to_excel(
                writer, sheet_name=rekap_sheet, index=False, header=False
            )
            writer.sheets[rekap_sheet].cell(row=1, column=1).font = Font(bold=True, size=14)

            rekap_rows = []
            for r in hasil:
                h = r["header"]
                akhir = r["validasi"][-1]["Status"]
                total_item = sum(b.get("subtotal") or 0 for b in r["barang"])
                rekap_rows.append({
                    "Nomor Seri": h.get("nomor_seri"),
                    "Tanggal": h.get("tanggal"),
                    "Nama Penjual": h.get("nama_penjual"),
                    "NPWP Penjual": h.get("npwp_penjual"),
                    "Nama Pembeli": h.get("nama_pembeli"),
                    "NPWP Pembeli": h.get("npwp_pembeli"),
                    "Jumlah Item": len(r["barang"]),
                    "Total Item": total_item,
                    "DPP": h.get("dpp"),
                    "PPN": h.get("ppn"),
                    "PPnBM": h.get("ppnbm"),
                    "Harga Jual Total": h.get("harga_jual_total"),
                    "Status": akhir,
                    "Nama File": h.get("nama_file"),
                })
            df_rekap = pd.DataFrame(rekap_rows)

            total_row = pd.DataFrame([{
                "Nomor Seri": "TOTAL", "Tanggal": "", "Nama Penjual": "",
                "NPWP Penjual": "", "Nama Pembeli": "", "NPWP Pembeli": "",
                "Jumlah Item": df_rekap["Jumlah Item"].sum() if not df_rekap.empty else 0,
                "Total Item": df_rekap["Total Item"].sum() if not df_rekap.empty else 0,
                "DPP": df_rekap["DPP"].sum() if not df_rekap.empty else 0,
                "PPN": df_rekap["PPN"].sum() if not df_rekap.empty else 0,
                "PPnBM": df_rekap["PPnBM"].sum() if not df_rekap.empty else 0,
                "Harga Jual Total": df_rekap["Harga Jual Total"].sum() if not df_rekap.empty else 0,
                "Status": "", "Nama File": "",
            }])
            df_rekap_full = pd.concat([df_rekap, total_row], ignore_index=True)

            tulis_df(
                writer, df_rekap_full, rekap_sheet, startrow=2,
                kolom_angka=["Total Item", "DPP", "PPN", "PPnBM",
                             "Harga Jual Total", "Jumlah Item"],
            )
            ws = writer.sheets[rekap_sheet]
            for col in range(1, len(df_rekap_full.columns) + 1):
                ws.cell(row=2 + len(df_rekap_full), column=col).font = Font(bold=True)

            for r in hasil:
                h = r["header"]
                raw_name = h.get("nomor_seri") or h.get("nama_file", "Faktur")
                sheet = safe_sheet_name(str(raw_name), used_names)

                df_info = pd.DataFrame([
                    ["Nomor Seri", h.get("nomor_seri")],
                    ["Tanggal", h.get("tanggal")],
                    ["NPWP Penjual", h.get("npwp_penjual")],
                    ["Nama Penjual", h.get("nama_penjual")],
                    ["Alamat Penjual", h.get("alamat_penjual")],
                    ["NPWP Pembeli", h.get("npwp_pembeli")],
                    ["Nama Pembeli", h.get("nama_pembeli")],
                    ["Alamat Pembeli", h.get("alamat_pembeli")],
                    ["DPP", h.get("dpp")],
                    ["PPN", h.get("ppn")],
                    ["PPnBM", h.get("ppnbm")],
                    ["Harga Jual Total", h.get("harga_jual_total")],
                ], columns=["Field", "Nilai"])

                row = 0
                row = tulis_judul(writer, sheet, "HEADER FAKTUR", row)
                row = tulis_df(writer, df_info, sheet, startrow=row,
                               kolom_angka=["Nilai"])

                df_barang = pd.DataFrame(r["barang"])
                row += 1
                row = tulis_judul(writer, sheet, "DETAIL BARANG", row)
                row = tulis_df(writer, df_barang, sheet, startrow=row,
                               kolom_angka=["harga_satuan", "qty", "subtotal"])

                df_validasi = pd.DataFrame(r["validasi"])
                row += 1
                row = tulis_judul(writer, sheet, "VALIDASI", row)
                tulis_df(writer, df_validasi, sheet, startrow=row,
                         kolom_angka=["Selisih"])

                ws = writer.sheets[sheet]
                for col_cells in ws.columns:
                    max_len = 0
                    letter = col_cells[0].column_letter
                    for cell in col_cells:
                        if cell.value:
                            max_len = max(max_len, len(str(cell.value)))
                    ws.column_dimensions[letter].width = min(max_len + 2, 60)

        st.download_button(
            label="⬇️ Download Excel (Rekap + 1 sheet per faktur)",
            data=buffer.getvalue(),
            file_name="faktur_pajak_rekap.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

else:
    st.info("Silakan upload file PDF faktur pajak terlebih dahulu.")
