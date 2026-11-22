import re

import pdfplumber


# pdf_metnini_oku sonrası Arapça metin burada tutulur.
# sunum_olustur bu değeri okuyup sıfırlar (arayuz imzası değişmez).
son_arapca_metin = ""

# Son okunan PDF yolu (Arapça alanı görsel kırpmak için).
son_pdf_yolu = ""


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


HITAPLAR = (
    "Aziz Müminler!",
    "Aziz Müslümanlar!",
    "Kıymetli Müslümanlar!",
    "Kıymetli Müminler!",
    "Kıymetli Kardeşlerim!",
    "Değerli Müminler!",
    "Değerli Müslümanlar!",
    "Değerli Kardeşlerim!",
    "Muhterem Müslümanlar!",
)


def baslik_ve_hitabi_ayir(satir):
    """Aynı satırdaki hutbe başlığını hitaptan ayırır."""
    for hitap in HITAPLAR:
        konum = satir.find(hitap)
        if konum > 0:
            baslik = satir[:konum].strip()
            kalan = satir[konum:].strip()
            if baslik and kalan:
                return [baslik, kalan]
    return [satir]


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
        return [], []

    temiz_satirlar = []
    arapca_satirlar = []

    for satir in metin.splitlines():
        satir = " ".join(satir.split()).strip()

        if not satir:
            continue

        # Arapça satırları silme; ayrı listede topla.
        if arapca_iceriyor_mu(satir):
            arapca_satirlar.append(satir)
            continue

        if gereksiz_satir_mi(satir):
            continue

        satir = dipnot_numaralarini_temizle(satir)

        if satir:
            # Başlık ile hitap aynı satırdaysa ayrı tut.
            temiz_satirlar.extend(baslik_ve_hitabi_ayir(satir))

    return temiz_satirlar, arapca_satirlar


def pdf_metnini_oku(pdf_yolu):
    global son_arapca_metin, son_pdf_yolu

    son_pdf_yolu = str(pdf_yolu)

    tum_satirlar = []
    tum_arapca_satirlar = []

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

            turkce_satirlar, arapca_satirlar = sutunu_oku(sol_sutun)
            tum_satirlar.extend(turkce_satirlar)
            tum_arapca_satirlar.extend(arapca_satirlar)

            turkce_satirlar, arapca_satirlar = sutunu_oku(sag_sutun)
            tum_satirlar.extend(turkce_satirlar)
            tum_arapca_satirlar.extend(arapca_satirlar)

    tum_metin = "\n".join(tum_satirlar).strip()
    son_arapca_metin = "\n".join(tum_arapca_satirlar).strip()

    if not tum_metin:
        raise ValueError(
            "PDF içinden Türkçe hutbe metni okunamadı."
        )

    return tum_metin