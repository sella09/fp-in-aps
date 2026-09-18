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
# VALIDASI
# =========================================================
TOLERANSI = 1.0
TARIF_PPN = 0.11


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
            "Cek": f"Item {i}: {b.get('nama_barang')} ({b.get('part_no')})",
            "Nilai": f"{rupiah(hs)} × {q}",
            "Selisih": selisih,
            "Status": status,
            "Keterangan": ket,
        })

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

    if dpp is not None and ppn is not None and harga_jual is not None:
        hitung = round(dpp + ppn + ppnbm, 2)
        selisih = round(harga_jual - hitung, 2)
        status = "✅" if abs(selisih) <= TOLERANSI else "❌"
        ket = "OK" if status == "✅" else f"Selisih {rupiah(selisih)}"
    else:
        selisih, status, ket = None, "⚠️", "Tidak bisa dicek"
    rows.append({
        "Cek": "DPP + PPN + PPnBM vs Harga Jual",
        "Nilai": f"{rupiah(dpp)} + {rupiah(ppn)} + {rupiah(ppnbm)} vs {rupiah(harga_jual)}",
        "Selisih": selisih,
        "Status": status, "Keterangan": ket,
    })

    if dpp and ppn is not None and dpp > 0:
        harusnya = round(dpp * TARIF_PPN, 2)
        selisih = round(ppn - harusnya, 2)
        if abs(selisih) <= TOLERANSI:
            status, ket = "✅", "OK"
        elif abs(selisih) <= 2:
            status, ket = "⚠️", f"Pembulatan {rupiah(selisih)}"
        else:
            status, ket = "❌", f"Selisih {rupiah(selisih)} (harusnya {rupiah(harusnya)})"
    else:
        harusnya, selisih, status, ket = None, None, "⚠️", "Tidak bisa dicek"
    rows.append({
        "Cek": "Tarif PPN 11% × DPP",
        "Nilai": f"PPN {rupiah(ppn)} vs {rupiah(harusnya)}",
        "Selisih": selisih,
        "Status": status, "Keterangan": ket,
    })

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
            # ---------- SHEET REKAP ----------
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

            next_row = tulis_df(
                writer, df_rekap_full, rekap_sheet, startrow=2,
                kolom_angka=["Total Item", "DPP", "PPN", "PPnBM",
                             "Harga Jual Total", "Jumlah Item"],
            )
            ws = writer.sheets[rekap_sheet]
            for col in range(1, len(df_rekap_full.columns) + 1):
                ws.cell(row=2 + len(df_rekap_full), column=col).font = Font(bold=True)

            # ---------- 1 SHEET PER FAKTUR ----------
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
