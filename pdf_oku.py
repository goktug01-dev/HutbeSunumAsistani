import re

import pdfplumber


def arapca_harf_mi(karakter):
    return (
        "\u0600" <= karakter <= "\u06ff"
        or "\u0750" <= karakter <= "\u077f"
        or "\u08a0" <= karakter <= "\u08ff"
        or "\ufb50" <= karakter <= "\ufdff"
        or "\ufe70" <= karakter <= "\ufeff"
    )


def arapca_iceriyor_mu(metin):
    return any(arapca_harf_mi(karakter) for karakter in metin)


def gereksiz_satir_mi(satir):
    satir = satir.strip()

    if not satir:
        return True

    kucuk = satir.lower()

    # Tarih ve kurum altbilgileri
    if kucuk.startswith("tarih:"):
        return True

    if "din hizmetleri genel müdürlüğü" in kucuk:
        return True

    if "diyanet işleri başkanlığı" in kucuk:
        return True

    # Arapça içeren satırlar
    if arapca_iceriyor_mu(satir):
        return True

    # Tek başına dipnot veya sayfa numarası
    if re.fullmatch(r"\d{1,3}", satir):
        return True

    # Kaynakça örnekleri:
    # 1 Târık, 86/13,14.
    # 3 Hicr, 15/9.
    # 5 Hâkim, Müstedrek, 3/144.
    if re.match(
        r"^\s*\d{1,3}\s+.+,\s*\d+\s*/\s*\d+",
        satir,
    ):
        return True

    return False


def dipnot_numaralarini_temizle(metin):
    # s.a.s. gibi ifadelerin sonuna yapışan dipnotlar
    metin = re.sub(
        r"(\(s\.a\.s\.\))\s*\d{1,3}(?=[’'])",
        r"\1",
        metin,
        flags=re.IGNORECASE,
    )

    # Cümle sonuna yapışan dipnotlar:
    # değildir.1  -> değildir.
    # biziz”3     -> biziz”
    # müjdeler.”2 -> müjdeler.”
    metin = re.sub(
        r'([.!?…”"]) *\d{1,3}(?=\s|$)',
        r"\1",
        metin,
    )

    return metin.strip()


def sutunu_oku(sutun):
    metin = sutun.extract_text(
        x_tolerance=2,
        y_tolerance=3,
        layout=False,
    )

    if not metin:
        return []

    temiz_satirlar = []

    for satir in metin.splitlines():
        satir = " ".join(satir.split()).strip()

        if gereksiz_satir_mi(satir):
            continue

        satir = dipnot_numaralarini_temizle(satir)

        if satir:
            temiz_satirlar.append(satir)

    return temiz_satirlar


def pdf_metnini_oku(pdf_yolu):
    tum_satirlar = []

    with pdfplumber.open(pdf_yolu) as pdf:
        for sayfa in pdf.pages:
            orta_nokta = sayfa.width / 2

            # Önce sol sütunun tamamını oku.
            sol_sutun = sayfa.crop(
                (
                    0,
                    0,
                    orta_nokta,
                    sayfa.height,
                )
            )

            # Sonra sağ sütunun tamamını oku.
            sag_sutun = sayfa.crop(
                (
                    orta_nokta,
                    0,
                    sayfa.width,
                    sayfa.height,
                )
            )

            tum_satirlar.extend(sutunu_oku(sol_sutun))
            tum_satirlar.extend(sutunu_oku(sag_sutun))

    tum_metin = "\n".join(tum_satirlar).strip()

    if not tum_metin:
        raise ValueError(
            "PDF içinden Türkçe hutbe metni okunamadı."
        )

    return tum_metin