import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path

from pdf_oku import pdf_metnini_oku
from ppt_olustur import sunum_olustur


# ─────────────────────────────────────────────
# RENKLER
# ─────────────────────────────────────────────

BEYAZ = "#FFFDF7"
KIRIK_BEYAZ = "#F8F3E7"
ALTIN = "#B68A2C"
ACIK_ALTIN = "#E3C978"
KOYU_ALTIN = "#80601E"
KOYU_YAZI = "#2E2A23"
GRI_YAZI = "#746E63"
YESIL = "#41644A"
KIRMIZI = "#A64040"

secilen_pdf_yolu = ""


# ─────────────────────────────────────────────
# YARDIMCI FONKSİYONLAR
# ─────────────────────────────────────────────

def dosya_adini_kisalt(dosya_yolu, en_fazla=58):
    dosya_adi = Path(dosya_yolu).name

    if len(dosya_adi) <= en_fazla:
        return dosya_adi

    return dosya_adi[:en_fazla - 3] + "..."


def durum_guncelle(metin, renk=GRI_YAZI):
    durum_etiketi.config(
        text=metin,
        fg=renk
    )

    pencere.update_idletasks()


def buton_uzerine_gel(event):
    event.widget.config(
        bg=KOYU_ALTIN,
        fg=BEYAZ
    )


def butondan_ayril(event):
    event.widget.config(
        bg=ALTIN,
        fg=BEYAZ
    )


def alt_buton_uzerine_gel(event):
    event.widget.config(
        bg=ACIK_ALTIN,
        fg=KOYU_YAZI
    )


def alt_butondan_ayril(event):
    event.widget.config(
        bg=KIRIK_BEYAZ,
        fg=KOYU_ALTIN
    )


# ─────────────────────────────────────────────
# PDF İŞLEMLERİ
# ─────────────────────────────────────────────

def pdf_sec():
    global secilen_pdf_yolu

    dosya_yolu = filedialog.askopenfilename(
        title="Hutbe PDF dosyasını seç",
        filetypes=[
            ("PDF dosyaları", "*.pdf"),
            ("Tüm dosyalar", "*.*"),
        ],
    )

    if not dosya_yolu:
        return

    secilen_pdf_yolu = dosya_yolu

    dosya_etiketi.config(
        text=f"✓  {dosya_adini_kisalt(dosya_yolu)}",
        fg=YESIL
    )

    pdf_olustur_butonu.config(
        state=tk.NORMAL,
        cursor="hand2"
    )

    durum_guncelle(
        "PDF seçildi. Sunumu oluşturmaya hazırsınız.",
        YESIL
    )


def pdfden_sunum_olustur():
    if not secilen_pdf_yolu:
        messagebox.showwarning(
            "PDF seçilmedi",
            "Lütfen önce bir hutbe PDF dosyası seçin."
        )
        return

    try:
        pdf_olustur_butonu.config(
            state=tk.DISABLED,
            text="SUNUM OLUŞTURULUYOR..."
        )

        durum_guncelle(
            "PDF okunuyor ve sunum hazırlanıyor...",
            ALTIN
        )

        pdf_metni = pdf_metnini_oku(secilen_pdf_yolu)
        dosya_yolu = sunum_olustur(pdf_metni)

        durum_guncelle(
            "Sunum başarıyla oluşturuldu.",
            YESIL
        )

        messagebox.showinfo(
            "Sunum hazır",
            "Hutbe sunumu başarıyla oluşturuldu.\n\n"
            f"Kaydedilen konum:\n{dosya_yolu}"
        )

    except Exception as hata:
        durum_guncelle(
            "Sunum oluşturulurken bir hata meydana geldi.",
            KIRMIZI
        )

        messagebox.showerror(
            "Hata oluştu",
            str(hata)
        )

    finally:
        pdf_olustur_butonu.config(
            state=tk.NORMAL,
            text="PDF'DEN SUNUM OLUŞTUR"
        )


# ─────────────────────────────────────────────
# YAPIŞTIRILAN METİN İŞLEMLERİ
# ─────────────────────────────────────────────

def yapistirilan_metinden_sunum_olustur():
    metin = metin_kutusu.get(
        "1.0",
        tk.END
    ).strip()

    if not metin:
        messagebox.showwarning(
            "Metin bulunamadı",
            "Lütfen hutbe metnini kutuya yapıştırın."
        )
        return

    try:
        metin_olustur_butonu.config(
            state=tk.DISABLED,
            text="SUNUM OLUŞTURULUYOR..."
        )

        durum_guncelle(
            "Yapıştırılan metin işleniyor...",
            ALTIN
        )

        dosya_yolu = sunum_olustur(metin)

        durum_guncelle(
            "Sunum başarıyla oluşturuldu.",
            YESIL
        )

        messagebox.showinfo(
            "Sunum hazır",
            "Hutbe sunumu başarıyla oluşturuldu.\n\n"
            f"Kaydedilen konum:\n{dosya_yolu}"
        )

    except Exception as hata:
        durum_guncelle(
            "Sunum oluşturulurken bir hata meydana geldi.",
            KIRMIZI
        )

        messagebox.showerror(
            "Hata oluştu",
            str(hata)
        )

    finally:
        metin_olustur_butonu.config(
            state=tk.NORMAL,
            text="YAPIŞTIRILAN METİNDEN SUNUM OLUŞTUR"
        )


def metin_kutusunu_temizle():
    metin_kutusu.delete(
        "1.0",
        tk.END
    )

    durum_guncelle(
        "Metin alanı temizlendi.",
        GRI_YAZI
    )


# ─────────────────────────────────────────────
# PENCERE
# ─────────────────────────────────────────────

pencere = tk.Tk()
pencere.title("Hutbe Sunum Asistanı")
pencere.geometry("960x780")
pencere.minsize(900, 720)
pencere.configure(bg=BEYAZ)

# Pencereyi ekranın ortasında aç.
pencere.update_idletasks()

ekran_genisligi = pencere.winfo_screenwidth()
ekran_yuksekligi = pencere.winfo_screenheight()

pencere_genisligi = 960
pencere_yuksekligi = 780

x_konumu = (
    ekran_genisligi - pencere_genisligi
) // 2

y_konumu = (
    ekran_yuksekligi - pencere_yuksekligi
) // 2

pencere.geometry(
    f"{pencere_genisligi}x{pencere_yuksekligi}"
    f"+{x_konumu}+{y_konumu}"
)


# ─────────────────────────────────────────────
# ANA ÇERÇEVE
# ─────────────────────────────────────────────

ana_cerceve = tk.Frame(
    pencere,
    bg=BEYAZ,
    highlightbackground=ALTIN,
    highlightthickness=2
)

ana_cerceve.pack(
    fill="both",
    expand=True,
    padx=22,
    pady=22
)


# ─────────────────────────────────────────────
# ÜST BAŞLIK ALANI
# ─────────────────────────────────────────────

ust_alan = tk.Frame(
    ana_cerceve,
    bg=BEYAZ
)

ust_alan.pack(
    fill="x",
    padx=35,
    pady=(22, 10)
)

ay_yildiz = tk.Label(
    ust_alan,
    text="☾ ★",
    font=("Georgia", 32, "bold"),
    fg=ALTIN,
    bg=BEYAZ
)

ay_yildiz.pack()

baslik = tk.Label(
    ust_alan,
    text="HUTBE SUNUM ASİSTANI",
    font=("Georgia", 24, "bold"),
    fg=KOYU_ALTIN,
    bg=BEYAZ
)

baslik.pack(
    pady=(3, 4)
)

alt_baslik = tk.Label(
    ust_alan,
    text="Hutbe metinlerini tek tıkla PowerPoint sunumuna dönüştürün",
    font=("Arial", 11),
    fg=GRI_YAZI,
    bg=BEYAZ
)

alt_baslik.pack()

altin_cizgi = tk.Frame(
    ust_alan,
    height=2,
    bg=ACIK_ALTIN
)

altin_cizgi.pack(
    fill="x",
    pady=(18, 0)
)


# ─────────────────────────────────────────────
# PDF ALANI
# ─────────────────────────────────────────────

pdf_cerceve = tk.Frame(
    ana_cerceve,
    bg=KIRIK_BEYAZ,
    highlightbackground=ACIK_ALTIN,
    highlightthickness=1
)

pdf_cerceve.pack(
    fill="x",
    padx=55,
    pady=(12, 15)
)

pdf_baslik = tk.Label(
    pdf_cerceve,
    text="PDF DOSYASINDAN SUNUM OLUŞTUR",
    font=("Georgia", 13, "bold"),
    fg=KOYU_ALTIN,
    bg=KIRIK_BEYAZ
)

pdf_baslik.pack(
    pady=(18, 4)
)

pdf_aciklama = tk.Label(
    pdf_cerceve,
    text=(
        "TDV hutbe PDF dosyanızı seçin. "
        "Arapça metinler ve dipnotlar otomatik olarak ayıklanır."
    ),
    font=("Arial", 10),
    fg=GRI_YAZI,
    bg=KIRIK_BEYAZ
)

pdf_aciklama.pack(
    pady=(0, 13)
)

pdf_sec_butonu = tk.Button(
    pdf_cerceve,
    text="📄  PDF DOSYASI SEÇ",
    font=("Arial", 11, "bold"),
    bg=KIRIK_BEYAZ,
    fg=KOYU_ALTIN,
    activebackground=ACIK_ALTIN,
    activeforeground=KOYU_YAZI,
    relief="solid",
    bd=1,
    padx=22,
    pady=10,
    cursor="hand2",
    command=pdf_sec
)

pdf_sec_butonu.pack()

pdf_sec_butonu.bind(
    "<Enter>",
    alt_buton_uzerine_gel
)

pdf_sec_butonu.bind(
    "<Leave>",
    alt_butondan_ayril
)

dosya_etiketi = tk.Label(
    pdf_cerceve,
    text="Henüz bir PDF dosyası seçilmedi.",
    font=("Arial", 9, "italic"),
    fg=GRI_YAZI,
    bg=KIRIK_BEYAZ,
    wraplength=750
)

dosya_etiketi.pack(
    pady=(10, 12)
)

pdf_olustur_butonu = tk.Button(
    pdf_cerceve,
    text="PDF'DEN SUNUM OLUŞTUR",
    font=("Arial", 12, "bold"),
    bg=ALTIN,
    fg=BEYAZ,
    activebackground=KOYU_ALTIN,
    activeforeground=BEYAZ,
    disabledforeground="#E7DDC5",
    relief="flat",
    bd=0,
    padx=30,
    pady=12,
    cursor="arrow",
    state=tk.DISABLED,
    command=pdfden_sunum_olustur
)

pdf_olustur_butonu.pack(
    pady=(0, 20)
)

pdf_olustur_butonu.bind(
    "<Enter>",
    buton_uzerine_gel
)

pdf_olustur_butonu.bind(
    "<Leave>",
    butondan_ayril
)


# ─────────────────────────────────────────────
# AYIRICI
# ─────────────────────────────────────────────

ayirici_alan = tk.Frame(
    ana_cerceve,
    bg=BEYAZ
)

ayirici_alan.pack(
    fill="x",
    padx=55,
    pady=2
)

sol_cizgi = tk.Frame(
    ayirici_alan,
    height=1,
    bg=ACIK_ALTIN
)

sol_cizgi.pack(
    side="left",
    fill="x",
    expand=True
)

veya_etiketi = tk.Label(
    ayirici_alan,
    text="  VEYA  ",
    font=("Georgia", 10, "bold"),
    fg=ALTIN,
    bg=BEYAZ
)

veya_etiketi.pack(
    side="left"
)

sag_cizgi = tk.Frame(
    ayirici_alan,
    height=1,
    bg=ACIK_ALTIN
)

sag_cizgi.pack(
    side="left",
    fill="x",
    expand=True
)


# ─────────────────────────────────────────────
# METİN ALANI
# ─────────────────────────────────────────────

metin_cerceve = tk.Frame(
    ana_cerceve,
    bg=BEYAZ
)

metin_cerceve.pack(
    fill="both",
    expand=True,
    padx=55,
    pady=(12, 8)
)

metin_baslik = tk.Label(
    metin_cerceve,
    text="HUTBE METNİNİ ELLE YAPIŞTIR",
    font=("Georgia", 12, "bold"),
    fg=KOYU_ALTIN,
    bg=BEYAZ
)

metin_baslik.pack(
    pady=(0, 8)
)

metin_kutusu_cercevesi = tk.Frame(
    metin_cerceve,
    bg=ACIK_ALTIN,
    padx=1,
    pady=1
)

metin_kutusu_cercevesi.pack(
    fill="both",
    expand=True
)

metin_kutusu = tk.Text(
    metin_kutusu_cercevesi,
    font=("Arial", 11),
    bg="#FFFFFF",
    fg=KOYU_YAZI,
    insertbackground=ALTIN,
    selectbackground=ACIK_ALTIN,
    selectforeground=KOYU_YAZI,
    relief="flat",
    bd=0,
    wrap="word",
    padx=14,
    pady=12,
    height=8
)

metin_kutusu.pack(
    fill="both",
    expand=True
)

metin_buton_alani = tk.Frame(
    metin_cerceve,
    bg=BEYAZ
)

metin_buton_alani.pack(
    pady=(10, 0)
)

metin_olustur_butonu = tk.Button(
    metin_buton_alani,
    text="YAPIŞTIRILAN METİNDEN SUNUM OLUŞTUR",
    font=("Arial", 10, "bold"),
    bg=ALTIN,
    fg=BEYAZ,
    activebackground=KOYU_ALTIN,
    activeforeground=BEYAZ,
    relief="flat",
    bd=0,
    padx=22,
    pady=10,
    cursor="hand2",
    command=yapistirilan_metinden_sunum_olustur
)

metin_olustur_butonu.pack(
    side="left",
    padx=5
)

metin_olustur_butonu.bind(
    "<Enter>",
    buton_uzerine_gel
)

metin_olustur_butonu.bind(
    "<Leave>",
    butondan_ayril
)

temizle_butonu = tk.Button(
    metin_buton_alani,
    text="TEMİZLE",
    font=("Arial", 10, "bold"),
    bg=KIRIK_BEYAZ,
    fg=KOYU_ALTIN,
    activebackground=ACIK_ALTIN,
    activeforeground=KOYU_YAZI,
    relief="solid",
    bd=1,
    padx=18,
    pady=9,
    cursor="hand2",
    command=metin_kutusunu_temizle
)

temizle_butonu.pack(
    side="left",
    padx=5
)

temizle_butonu.bind(
    "<Enter>",
    alt_buton_uzerine_gel
)

temizle_butonu.bind(
    "<Leave>",
    alt_butondan_ayril
)


# ─────────────────────────────────────────────
# ALT DURUM VE İMZA
# ─────────────────────────────────────────────

alt_alan = tk.Frame(
    ana_cerceve,
    bg=BEYAZ
)

alt_alan.pack(
    fill="x",
    padx=40,
    pady=(5, 14)
)

durum_etiketi = tk.Label(
    alt_alan,
    text="Sunum oluşturmak için PDF seçin veya hutbe metnini yapıştırın.",
    font=("Arial", 9),
    fg=GRI_YAZI,
    bg=BEYAZ
)

durum_etiketi.pack(
    pady=(2, 9)
)

imza_cizgisi = tk.Frame(
    alt_alan,
    height=1,
    bg=ACIK_ALTIN
)

imza_cizgisi.pack(
    fill="x",
    padx=120,
    pady=(0, 8)
)

imza = tk.Label(
    alt_alan,
    text="Bu program Uğur KOCATÜRK için özel olarak tasarlanmıştır.",
    font=("Georgia", 9, "italic"),
    fg=KOYU_ALTIN,
    bg=BEYAZ
)

imza.pack()

pencere.mainloop()