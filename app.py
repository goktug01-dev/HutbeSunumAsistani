"""
Hutbe Sunum Asistanı - Web sürümü (Streamlit)

Yerelde denemek için:   streamlit run app.py
"""

import html
import io
import math
import traceback
from datetime import datetime
from urllib.parse import quote

import streamlit as st
from pptx import Presentation

from pdf_oku import pdf_metnini_oku
from ppt_olustur import (
    TURKIYE_SAATI,
    baslik_mi,
    bu_haftanin_cumasi,
    cumlelere_ayir,
    sunum_olustur,
)


st.set_page_config(
    page_title="Hutbe Sunum Asistanı | Şehit Halit Yaşar Mine Camii",
    page_icon="🕌",
    layout="centered",
)


# ─────────────────────────────────────────────
# KİŞİSEL BİLGİLER
# ─────────────────────────────────────────────

CAMI_ADI = "Şehit Halit Yaşar Mine Camii"
HITAP = "Uğur Hocam"
IMZA = "Oğlun Göktuğ'dan, sevgiyle"


# ─────────────────────────────────────────────
# RENKLER (.streamlit/config.toml ile uyumlu)
# ─────────────────────────────────────────────

CAMI_YESILI = "#1F3D30"
FILDISI = "#FBF6EA"
ALTIN = "#B8913F"
KOYU_ALTIN = "#8A6A26"
MUREKKEP = "#2A241B"
SOLUK = "#6E6553"


# ─────────────────────────────────────────────
# SÜSLEME (SVG)
# ─────────────────────────────────────────────

def svg_adresi(svg):
    return "data:image/svg+xml;utf8," + quote(svg)


def kemer_svg():
    """Mihrap kemeri: iki kavisin tepede sivri birleştiği kemer."""
    return (
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 200 100" '
        'preserveAspectRatio="none">'
        f'<path d="M0 100 V62 C0 30 64 18 100 1 C136 18 200 30 200 62 V100" '
        f'fill="{FILDISI}" stroke="{ALTIN}" stroke-width="3" '
        'vector-effect="non-scaling-stroke"/>'
        "</svg>"
    )


def ay_yildiz_svg():
    merkez_x, merkez_y = 41, 24
    noktalar = []
    for i in range(10):
        yaricap = 7 if i % 2 == 0 else 2.9
        aci = math.radians(-90 + i * 36)
        noktalar.append(
            f"{merkez_x + yaricap * math.cos(aci):.2f},"
            f"{merkez_y + yaricap * math.sin(aci):.2f}"
        )

    return (
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 52 48">'
        f'<circle cx="21" cy="24" r="15" fill="{ALTIN}"/>'
        f'<circle cx="25.5" cy="24" r="12" fill="{FILDISI}"/>'
        f'<polygon points="{" ".join(noktalar)}" fill="{ALTIN}"/>'
        "</svg>"
    )


def yildiz_deseni_svg():
    """Arka plan dokusu: sekiz kollu yıldızlar ve onları saran baklavalar."""

    def sekiz_kollu(cx, cy, r):
        # {8/3} yıldız: her köşe 3 sonrakine bağlanır
        noktalar = [
            (
                cx + r * math.cos(math.radians(22.5 + k * 135)),
                cy + r * math.sin(math.radians(22.5 + k * 135)),
            )
            for k in range(8)
        ]
        return "M" + " L".join(f"{x:.1f} {y:.1f}" for x, y in noktalar) + " Z"

    yildizlar = " ".join(
        sekiz_kollu(x, y, 13)
        for x, y in [(36, 36), (0, 0), (72, 0), (0, 72), (72, 72)]
    )

    return (
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 72 72">'
        f'<g fill="none" stroke="{ALTIN}" stroke-width="0.8" opacity="0.13">'
        f'<path d="{yildizlar}"/>'
        '<path d="M0 36 L36 0 L72 36 L36 72 Z"/>'
        "</g></svg>"
    )


def cami_svg(cx, olcek, minare_boyu, zemin=280):
    """Tek bir caminin silüeti: kubbe, yarım kubbeler, minareler, ışıklı pencereler."""
    s = olcek
    kubbe_r = 70 * s
    govde_yarisi = 125 * s
    govde_ust = zemin - 70 * s
    kasnak_ust = govde_ust - 14 * s
    kubbe_tepe = kasnak_ust - kubbe_r * 1.05

    parcalar = [
        # Gövde ve kubbe kasnağı
        f'<rect x="{cx - govde_yarisi:.1f}" y="{govde_ust:.1f}" '
        f'width="{2 * govde_yarisi:.1f}" height="{zemin - govde_ust:.1f}"/>',
        f'<rect x="{cx - kubbe_r - 5 * s:.1f}" y="{kasnak_ust:.1f}" '
        f'width="{2 * (kubbe_r + 5 * s):.1f}" height="{14 * s + 1:.1f}"/>',
        # Ana kubbe
        f'<path d="M{cx - kubbe_r:.1f} {kasnak_ust:.1f} '
        f'A{kubbe_r:.1f} {kubbe_r * 1.05:.1f} 0 0 1 {cx + kubbe_r:.1f} {kasnak_ust:.1f} Z"/>',
    ]

    # Yarım kubbeler
    for yon in (-1, 1):
        yarim_cx = cx + yon * (kubbe_r + 22 * s)
        yarim_r = 32 * s
        parcalar.append(
            f'<path d="M{yarim_cx - yarim_r:.1f} {govde_ust + 1:.1f} '
            f'A{yarim_r:.1f} {yarim_r:.1f} 0 0 1 {yarim_cx + yarim_r:.1f} {govde_ust + 1:.1f} Z"/>'
        )

    alemler = [(cx, kubbe_tepe, 16 * s)]

    # Minareler (kalem minare: ince gövde, şerefeler, sivri külah)
    for yon in (-1, 1):
        mx = cx + yon * (govde_yarisi + 22 * s)
        genislik = 11 * s
        tepe = zemin - minare_boyu * s
        kulah_dibi = tepe + 34 * s
        parcalar.append(
            f'<rect x="{mx - genislik / 2:.1f}" y="{kulah_dibi:.1f}" '
            f'width="{genislik:.1f}" height="{zemin - kulah_dibi:.1f}"/>'
        )
        parcalar.append(
            f'<path d="M{mx - genislik / 2 - 1:.1f} {kulah_dibi + 1:.1f} '
            f'L{mx:.1f} {tepe:.1f} L{mx + genislik / 2 + 1:.1f} {kulah_dibi + 1:.1f} Z"/>'
        )
        for oran in (0.28, 0.5):
            serefe_y = kulah_dibi + (zemin - kulah_dibi) * oran
            parcalar.append(
                f'<rect x="{mx - genislik:.1f}" y="{serefe_y:.1f}" '
                f'width="{2 * genislik:.1f}" height="{4 * s:.1f}"/>'
            )
        alemler.append((mx, tepe, 9 * s))

    # Işıklı pencereler
    pencereler = []
    adim = 26 * s
    adet = int((2 * govde_yarisi - 30 * s) // adim)
    baslangic = cx - (adet - 1) * adim / 2
    for i in range(adet):
        px = baslangic + i * adim
        py = zemin - 42 * s
        g, y = 6 * s, 14 * s
        pencereler.append(
            f'<path d="M{px - g / 2:.1f} {py + y:.1f} V{py + g / 2:.1f} '
            f'A{g / 2:.1f} {g / 2:.1f} 0 0 1 {px + g / 2:.1f} {py + g / 2:.1f} '
            f'V{py + y:.1f} Z"/>'
        )

    # Altın alemler (direk + hilal)
    alem_ciz = []
    for ax, ay, boy in alemler:
        ust = ay - boy
        r = boy * 0.28
        alem_ciz.append(
            f'<line x1="{ax:.1f}" y1="{ay:.1f}" x2="{ax:.1f}" y2="{ust + r:.1f}"/>'
            f'<path d="M{ax - r:.1f} {ust:.1f} A{r:.1f} {r:.1f} 0 0 0 {ax + r:.1f} {ust:.1f}" fill="none"/>'
        )

    return (
        "".join(parcalar),
        "".join(pencereler),
        "".join(alem_ciz),
    )


def bizim_cami_svg(cx, olcek, zemin=280):
    """
    Şehit Halit Yaşar Mine Camii: tek minare, gece yeşil ışıkla aydınlanan
    minare gövdesi, ışıklı şerefe, basık ana kubbe ve yanında küçük kubbe.
    """
    s = olcek
    govde_sol, govde_sag = cx - 105 * s, cx + 130 * s
    govde_ust = zemin - 78 * s
    kubbe_cx, kubbe_r = cx + 18 * s, 64 * s
    kasnak_ust = govde_ust - 12 * s
    kubbe_tepe = kasnak_ust - kubbe_r * 0.82

    parcalar = [
        f'<rect x="{govde_sol:.1f}" y="{govde_ust:.1f}" '
        f'width="{govde_sag - govde_sol:.1f}" height="{zemin - govde_ust:.1f}"/>',
        f'<rect x="{kubbe_cx - kubbe_r - 4 * s:.1f}" y="{kasnak_ust:.1f}" '
        f'width="{2 * (kubbe_r + 4 * s):.1f}" height="{12 * s + 1:.1f}"/>',
        f'<path d="M{kubbe_cx - kubbe_r:.1f} {kasnak_ust:.1f} '
        f'A{kubbe_r:.1f} {kubbe_r * 0.82:.1f} 0 0 1 {kubbe_cx + kubbe_r:.1f} {kasnak_ust:.1f} Z"/>',
    ]

    kucuk_cx, kucuk_r = cx + 108 * s, 22 * s
    parcalar.append(
        f'<path d="M{kucuk_cx - kucuk_r:.1f} {govde_ust + 1:.1f} '
        f'A{kucuk_r:.1f} {kucuk_r * 0.9:.1f} 0 0 1 {kucuk_cx + kucuk_r:.1f} {govde_ust + 1:.1f} Z"/>'
    )

    # Tek minare
    mx = cx - 132 * s
    genislik = 13 * s
    tepe = zemin - 250 * s
    kulah_dibi = tepe + 40 * s
    serefe_y = kulah_dibi + 46 * s
    parcalar.append(
        f'<rect x="{mx - genislik / 2:.1f}" y="{kulah_dibi:.1f}" '
        f'width="{genislik:.1f}" height="{zemin - kulah_dibi:.1f}"/>'
        f'<path d="M{mx - genislik / 2 - 1:.1f} {kulah_dibi + 1:.1f} '
        f'L{mx:.1f} {tepe:.1f} L{mx + genislik / 2 + 1:.1f} {kulah_dibi + 1:.1f} Z"/>'
    )

    # Gece ışıkları: yeşil minare gövdesi ve şerefe halkası
    yesil_ust = serefe_y + 6 * s
    yesil_alt = zemin - 20 * s
    isiklar = (
        f'<rect x="{mx - genislik * 1.6:.1f}" y="{yesil_ust:.1f}" '
        f'width="{genislik * 3.2:.1f}" height="{yesil_alt - yesil_ust:.1f}" '
        'fill="#3FCB74" opacity="0.35" filter="url(#parilti)"/>'
        f'<rect x="{mx - genislik / 2:.1f}" y="{yesil_ust:.1f}" '
        f'width="{genislik:.1f}" height="{yesil_alt - yesil_ust:.1f}" fill="#2FA65C"/>'
        f'<rect x="{mx - genislik * 1.4:.1f}" y="{serefe_y - 3 * s:.1f}" '
        f'width="{genislik * 2.8:.1f}" height="{11 * s:.1f}" '
        'fill="#F2A93B" opacity="0.6" filter="url(#parilti)"/>'
        f'<rect x="{mx - genislik:.1f}" y="{serefe_y:.1f}" '
        f'width="{2 * genislik:.1f}" height="{5 * s:.1f}" fill="#E9A23B"/>'
    )

    # Büyük kemerli pencereler
    pencereler = []
    for i in range(6):
        px = govde_sol + 30 * s + i * 34 * s
        py = zemin - 58 * s
        g, y = 14 * s, 30 * s
        pencereler.append(
            f'<path d="M{px - g / 2:.1f} {py + y:.1f} V{py + g / 2:.1f} '
            f'A{g / 2:.1f} {g / 2:.1f} 0 0 1 {px + g / 2:.1f} {py + g / 2:.1f} '
            f'V{py + y:.1f} Z"/>'
        )

    alem_ciz = []
    for ax, ay, boy in [(kubbe_cx, kubbe_tepe, 16 * s), (mx, tepe, 10 * s)]:
        ust = ay - boy
        r = boy * 0.28
        alem_ciz.append(
            f'<line x1="{ax:.1f}" y1="{ay:.1f}" x2="{ax:.1f}" y2="{ust + r:.1f}"/>'
            f'<path d="M{ax - r:.1f} {ust:.1f} A{r:.1f} {r:.1f} 0 0 0 {ax + r:.1f} {ust:.1f}" fill="none"/>'
        )

    return "".join(parcalar), "".join(pencereler), "".join(alem_ciz), isiklar


# Bizim cami silüette bu noktada durur (telefonda ekranın ortasına bu gelir).
BIZIM_CAMI_X = 700


def siluet_svg():
    """Alt kenarda, mihrap panosunun iki yanından görünen camiler.
    Orta boş bırakıldı: pano ve imza orada duruyor."""
    bizim = bizim_cami_svg(BIZIM_CAMI_X, 0.95)
    camiler = [
        bizim[:3],
        cami_svg(1700, 0.85, 205),
        cami_svg(250, 0.6, 165),
        cami_svg(2150, 0.6, 165),
    ]
    govdeler = "".join(c[0] for c in camiler)
    pencereler = "".join(c[1] for c in camiler)
    alemler = "".join(c[2] for c in camiler)

    return (
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 2400 280" '
        'preserveAspectRatio="xMidYMax meet">'
        '<defs><filter id="parilti" x="-100%" y="-100%" width="300%" height="300%">'
        '<feGaussianBlur stdDeviation="5"/></filter></defs>'
        f'<g fill="#142A20">{govdeler}<rect x="0" y="268" width="2400" height="12"/></g>'
        f'<g fill="{ALTIN}" opacity="0.55">{pencereler}</g>'
        f"{bizim[3]}"
        f'<g stroke="{ALTIN}" stroke-width="2.2" stroke-linecap="round">{alemler}</g>'
        "</svg>"
    )


# ─────────────────────────────────────────────
# STİL
# ─────────────────────────────────────────────

SILUET_ADRESI = svg_adresi(siluet_svg())
GOK_KATMANLARI = (
    "radial-gradient(ellipse 50rem 32rem at 50% 6rem, rgba(184, 145, 63, 0.22), transparent 70%),"
    f'url("{svg_adresi(yildiz_deseni_svg())}") center top / 4.5rem repeat,'
    f"{CAMI_YESILI}"
)
# Telefonda bizim caminin ekranın ortasına gelmesi için kaydırma payı
TELEFON_KAYDIRMA = f"{(1200 - BIZIM_CAMI_X) / 280 * 9:.2f}rem"

st.markdown(
    f"""
<style>
/* Streamlit'in kendi üst çubuğu ve alt yazısı gizlensin */
header[data-testid="stHeader"], footer, #MainMenu {{
    display: none !important;
}}

/* Gece gökyüzü: kandil ışıltısı + yıldız deseni + alt kenarda camiler */
.stApp {{
    background:
        url("{SILUET_ADRESI}") bottom center / auto clamp(9rem, 20vw, 16rem) no-repeat fixed,
        {GOK_KATMANLARI};
}}

.siluet-telefon {{
    display: none;
}}

[data-testid="stMainBlockContainer"] {{
    max-width: 40rem;
    padding: clamp(7.5rem, 32vw, 14rem) 1rem 2.5rem;
}}

/* Dar ekranda pano camileri kapatmasın: en altta camilere yer açılsın */
/* Telefonda camiler sayfanın en sonunda, imzanın altında dursun;
   ortaya da bizim cami gelsin */
@media (max-width: 52rem) {{
    .stApp {{
        background: {GOK_KATMANLARI};
    }}

    [data-testid="stMainBlockContainer"] {{
        padding-bottom: 0;
    }}

    .siluet-telefon {{
        display: block;
        height: 9rem;
        margin: 1rem -1rem 0;
        background: url("{SILUET_ADRESI}") calc(50% + {TELEFON_KAYDIRMA}) bottom / auto 9rem no-repeat;
    }}
}}

/* ── Mihrap panosu ── */
.st-key-mihrap {{
    position: relative;
    background: {FILDISI};
    border: 1.5px solid {ALTIN};
    border-top: none;
    padding: 0 clamp(1.25rem, 5vw, 2.75rem) 2.25rem;
    gap: 1.25rem;
}}

.st-key-mihrap::before {{
    content: "";
    position: absolute;
    left: -1.5px;
    right: -1.5px;
    bottom: 100%;
    height: clamp(6rem, 28vw, 11.5rem);
    background: url("{svg_adresi(kemer_svg())}") no-repeat bottom / 100% 100%;
}}

.st-key-mihrap::after {{
    content: "";
    position: absolute;
    left: 50%;
    bottom: calc(100% + clamp(1.25rem, 7vw, 3rem));
    width: 3.25rem;
    height: 3rem;
    transform: translateX(-50%);
    background: url("{svg_adresi(ay_yildiz_svg())}") no-repeat center / contain;
}}

.baslik {{
    text-align: center;
    margin: 0;
}}

/* Caminin adı: kapı üstündeki kitabe gibi, iki yanında altın çizgi */
.kitabe {{
    display: flex;
    align-items: center;
    gap: 0.75rem;
    color: {KOYU_ALTIN};
    font-family: "Alegreya", serif;
    font-style: italic;
    font-weight: 500;
    font-size: clamp(1rem, 3.6vw, 1.15rem);
    margin: 0 0 1.25rem;
    white-space: nowrap;
}}

.kitabe::before,
.kitabe::after {{
    content: "";
    flex: 1;
    height: 1px;
    background: linear-gradient(90deg, transparent, {ALTIN});
}}

.kitabe::after {{
    background: linear-gradient(90deg, {ALTIN}, transparent);
}}

.baslik .selam {{
    font-family: "Alegreya", serif;
    font-style: italic;
    font-size: clamp(1.25rem, 4.5vw, 1.5rem);
    color: {SOLUK};
    margin: 0;
    max-width: none;
}}

.baslik h1 {{
    font-size: clamp(2.4rem, 9vw, 3.25rem);
    line-height: 1.05;
    color: {CAMI_YESILI};
    padding: 0;
    margin: 0.1rem 0 0.9rem;
}}

.cuma {{
    display: inline-block;
    background: #F1E4C3;
    color: {KOYU_ALTIN};
    font-weight: 600;
    font-size: 0.95rem;
    border-radius: 999px;
    padding: 0.3rem 0.95rem;
    margin: 0 0 0.75rem;
}}

.baslik [data-testid="stHeaderActionElements"] {{
    display: none;
}}

.baslik p {{
    color: {SOLUK};
    font-size: 1.05rem;
    line-height: 1.5;
    max-width: 26rem;
    margin: 0 auto;
    text-wrap: balance;
}}

/* ── PDF yükleme alanı ── */
[data-testid="stFileUploaderDropzone"] {{
    display: flex !important;
    flex-direction: column !important;
    align-items: center !important;
    justify-content: center;
    background: #F5ECD7;
    border: 1.5px dashed {ALTIN};
    border-radius: 0.75rem;
    padding: 2rem 1.25rem;
}}

/* Dosya seçildikten sonra kutu küçülsün */
[data-testid="stFileUploaderDropzone"]:has([data-testid="stFileChips"]) {{
    border-style: solid;
    padding: 0.75rem;
}}

/* İngilizce yazılar ve tek dosyada anlamsız "+" butonu gizlensin */
[data-testid="stFileUploaderDropzoneInstructions"],
[data-testid="stFileUploaderDropzone"] [data-testid="stBaseButton-borderlessIcon"] {{
    display: none !important;
}}

[data-testid="stFileUploaderDropzone"] [data-testid="stBaseButton-secondary"] {{
    min-height: 3.25rem;
    padding: 0 1.75rem;
    background: {FILDISI};
    border: 1.5px solid {ALTIN};
    color: {CAMI_YESILI};
}}

[data-testid="stFileUploaderDropzone"] [data-testid="stBaseButton-secondary"]:hover {{
    background: {CAMI_YESILI};
    border-color: {CAMI_YESILI};
    color: {FILDISI};
}}

[data-testid="stFileUploaderDropzone"] [data-testid="stBaseButton-secondary"] p {{
    font-size: 0 !important;
}}

[data-testid="stFileUploaderDropzone"] [data-testid="stBaseButton-secondary"] p::after {{
    content: "Hutbe PDF'ini seç";
    font-size: 1.1rem;
    font-weight: 600;
}}

/* ── Butonlar ── */
.st-key-mihrap button {{
    min-height: 3.25rem;
    font-size: 1.1rem;
    font-weight: 600;
}}

/* ── Sonuç ── */
.sonuc {{
    text-align: center;
    border-top: 1px solid #E4D3A8;
    padding-top: 1.5rem;
}}

.sonuc .hutbe-adi {{
    font-family: "Alegreya", serif;
    font-size: 1.6rem;
    font-weight: 700;
    color: {MUREKKEP};
    margin: 0;
    line-height: 1.2;
}}

.sonuc .slayt-sayisi {{
    color: {KOYU_ALTIN};
    margin: 0.35rem 0 0;
}}

.dua {{
    text-align: center;
    color: {KOYU_ALTIN};
    font-family: "Alegreya", serif;
    font-style: italic;
    font-size: 1.15rem;
    margin: -0.25rem 0 0;
}}

/* ── Alt imza (yeşil zeminde) ── */
.imza {{
    text-align: center;
    margin-top: 1.75rem;
    line-height: 1.4;
}}

.imza .uygulama {{
    color: #C9B98F;
    font-family: "Alegreya", serif;
    font-weight: 700;
    font-size: 1.05rem;
    margin: 0;
}}

.imza .not-oglundan {{
    color: #A99C78;
    font-family: "Alegreya", serif;
    font-style: italic;
    margin: 0;
}}
</style>
""",
    unsafe_allow_html=True,
)


# ─────────────────────────────────────────────
# YARDIMCI FONKSİYONLAR
# ─────────────────────────────────────────────

def turkce_kucuk(metin):
    return metin.replace("I", "ı").replace("İ", "i").lower()


def turkce_buyuk(metin):
    return metin.replace("i", "İ").replace("ı", "I").upper()


def baslik_duzenle(metin):
    """'HELAL HARAM DUYARLILIĞI' -> 'Helal Haram Duyarlılığı'"""
    kelimeler = turkce_kucuk(metin).split()
    return " ".join(
        turkce_buyuk(kelime[0]) + kelime[1:] for kelime in kelimeler
    )


def hutbe_basligini_bul(metin):
    for cumle in cumlelere_ayir(metin)[:3]:
        if baslik_mi(cumle):
            return baslik_duzenle(cumle)

    return "Hutbe sunumu"


AYLAR = (
    "Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran",
    "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık",
)


def selam_metni(simdi):
    """Güne ve saate göre selam (Türkiye saati)."""
    if simdi.weekday() == 4:
        return "Hayırlı Cumalar,"

    if simdi.weekday() == 3 and simdi.hour >= 17:
        return "Cuma geceniz mübarek olsun,"

    if 5 <= simdi.hour < 11:
        return "Hayırlı sabahlar,"

    if 11 <= simdi.hour < 17:
        return "Hayırlı günler,"

    return "Hayırlı akşamlar,"


def cuma_metni(simdi):
    cuma = bu_haftanin_cumasi()
    kalan_gun = (cuma.date() - simdi.date()).days
    tarih = f"{cuma.day} {AYLAR[cuma.month - 1]}"

    if kalan_gun == 0:
        return f"Bugün Cuma, {tarih}"

    if kalan_gun == 1:
        return f"Yarın Cuma, {tarih}"

    return f"Cumaya {kalan_gun} gün var, {tarih}"


def dosya_adi():
    tarih = bu_haftanin_cumasi().strftime("%d-%m-%Y")
    return f"{tarih}_HutbeSunumu.pptx"


def sunumu_hazirla(metin, kaynak):
    cikti = io.BytesIO()
    sunum_olustur(metin, cikti)

    veri = cikti.getvalue()
    slayt_sayisi = len(Presentation(io.BytesIO(veri)).slides)

    st.session_state["sonuc"] = {
        "kaynak": kaynak,
        "veri": veri,
        "dosya_adi": dosya_adi(),
        "baslik": hutbe_basligini_bul(metin),
        "slayt_sayisi": slayt_sayisi,
    }


def hatayi_goster(hata, ipucu):
    st.session_state.pop("sonuc", None)
    st.error(f"{hata}\n\n{ipucu}")

    with st.expander("Teknik ayrıntı (Göktuğ'a gönderin)"):
        st.code(traceback.format_exc(), language=None)


# ─────────────────────────────────────────────
# SAYFA
# ─────────────────────────────────────────────

with st.container(key="mihrap"):
    simdi = datetime.now(TURKIYE_SAATI)

    st.markdown(
        f"""
<div class="baslik">
  <div class="kitabe">{CAMI_ADI}</div>
  <p class="selam">{selam_metni(simdi)}</p>
  <h1>{HITAP}</h1>
  <span class="cuma">{cuma_metni(simdi)}</span>
  <p>Hutbe PDF'ini seçin, sunumunuz birkaç saniyede hazırlansın.</p>
</div>
""",
        unsafe_allow_html=True,
    )

    pdf = st.file_uploader(
        "Hutbe PDF'i",
        type=["pdf"],
        label_visibility="collapsed",
    )

    sonuc = st.session_state.get("sonuc")

    if pdf is None and sonuc and sonuc["kaynak"].startswith("pdf:"):
        st.session_state.pop("sonuc", None)

    elif pdf is not None and (
        not sonuc or sonuc["kaynak"] != f"pdf:{pdf.file_id}"
    ):
        with st.spinner("Sunum hazırlanıyor..."):
            try:
                sunumu_hazirla(pdf_metnini_oku(pdf), f"pdf:{pdf.file_id}")
            except Exception as hata:
                hatayi_goster(
                    hata,
                    "Diyanet'in yayımladığı hutbe PDF'ini seçtiğinizden "
                    "emin olun.",
                )

    sonuc = st.session_state.get("sonuc")

    if sonuc:
        st.markdown(
            f"""
<div class="sonuc">
  <p class="hutbe-adi">{html.escape(sonuc["baslik"])}</p>
  <p class="slayt-sayisi">{sonuc["slayt_sayisi"]} slayt hazır</p>
</div>
""",
            unsafe_allow_html=True,
        )

        st.download_button(
            "Sunumu indir",
            data=sonuc["veri"],
            file_name=sonuc["dosya_adi"],
            mime=(
                "application/vnd.openxmlformats-officedocument"
                ".presentationml.presentation"
            ),
            type="primary",
            icon=":material/download:",
            width="stretch",
        )

        st.markdown(
            '<p class="dua">Rabbim emeğinizi kabul etsin hocam.</p>',
            unsafe_allow_html=True,
        )

    with st.expander("PDF yok mu? Metni yapıştırın"):
        metin = st.text_area(
            "Hutbe metni",
            height=220,
            label_visibility="collapsed",
            placeholder="Hutbe metnini buraya yapıştırın",
        )

        if st.button(
            "Metinden sunum oluştur",
            width="stretch",
            disabled=not metin.strip(),
        ):
            with st.spinner("Sunum hazırlanıyor..."):
                try:
                    sunumu_hazirla(metin.strip(), "metin")
                    st.rerun()
                except Exception as hata:
                    hatayi_goster(
                        hata,
                        "Metnin tamamını kopyaladığınızdan emin olun.",
                    )

st.markdown(
    f"""
<div class="imza">
  <p class="uygulama">Hutbe Sunum Asistanı</p>
  <p class="not-oglundan">{IMZA}</p>
</div>
<div class="siluet-telefon"></div>
""",
    unsafe_allow_html=True,
)
