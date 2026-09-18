def ekstrak_barang(teks):
    items = []
    posisi_part = [(m.start(), m) for m in re.finditer(r"-\s*Part No\s*:\s*(\S+)", teks)]

    if posisi_part:
        # === FORMAT A: dengan "- Part No" ===
        m_end = re.search(r"Harga Jual / Penggantian", teks)
        batas_akhir = m_end.start() if m_end else len(teks)

        for i, (start, m) in enumerate(posisi_part):
            end = posisi_part[i + 1][0] if i + 1 < len(posisi_part) else batas_akhir
            blok = teks[start:end]
            part_no = m.group(1).strip()

            before = teks[max(0, start - 120):start]
            m_nama = re.search(r"(?:^|\s)\d{1,2}\s+\d{6}\s+(.+?)\s*$", before)
            nama = m_nama.group(1).strip() if m_nama else ""

            m2 = re.search(r"Rp\s*([\d.,]+)\s*x\s*([\d.,]+)\s*(\w+)", blok)
            harga = parse_angka(m2.group(1)) if m2 else None
            qty = parse_angka(m2.group(2)) if m2 else None
            satuan = m2.group(3).strip() if m2 else None

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
                "nama_barang": nama, "part_no": part_no,
                "harga_satuan": harga, "qty": qty,
                "satuan": satuan, "subtotal": subtotal,
            })
        return items

    # === FORMAT B: tanpa "- Part No" (NADE / Sinergi style) ===
    # Pola: No Kode NamaBarang (mungkin ada tanda kurung) Rp harga x qty Satuan
    #       ... Potongan Harga ... PPnBM ... SUBTOTAL
    # Cari setiap blok item: "No Kode" diikuti sampai sebelum "No Kode" berikutnya

    # Cari posisi awal tiap item: "\d{1,3} \d{6} NamaBarang ... Rp"
    pola_awal = re.compile(
        r"(?:^|\s)(?P<no>\d{1,3})\s+(?P<kode>\d{6})\s+(?P<nama>[A-Za-z][^\n]*?)\s+Rp\s*(?P<harga>[\d.,]+)\s*x\s*(?P<qty>[\d.,]+)\s*(?P<satuan>\w+)",
        re.DOTALL
    )
    matches = list(pola_awal.finditer(teks))

    for i, m in enumerate(matches):
        start = m.end()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(teks)
        # batasi blok agar tidak menyerobot footer
        blok_lanjutan = teks[start:end]

        # Cari subtotal: angka besar terakhir di blok (setelah PPnBM/PnPbM)
        m_sub = re.search(
            r"[Pp]+[Nn]?[Bb]?[Mm]?\s*\(\s*0[,.]00%\s*\)\s*=\s*Rp\s*[\d.,]+\s*([\d.,]+)",
            blok_lanjutan
        )
        if m_sub:
            subtotal = parse_angka(m_sub.group(1))
        else:
            angka = re.findall(r"(\d{1,3}(?:\.\d{3})+(?:,\d{2})?)", blok_lanjutan)
            subtotal = parse_angka(angka[-1]) if angka else None

        items.append({
            "nama_barang": m.group("nama").strip(),
            "part_no": m.group("kode").strip(),
            "harga_satuan": parse_angka(m.group("harga")),
            "qty": parse_angka(m.group("qty")),
            "satuan": m.group("satuan").strip(),
            "subtotal": subtotal,
        })
    return items
