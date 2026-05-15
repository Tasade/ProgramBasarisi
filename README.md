# Balıkesir Üniversitesi — Program Başarı Sırası Paneli

Tek dosya Streamlit uygulaması.

## 📁 Dosya Yapısı

```
repo/
├── app.py                                          # Tüm kod (tek dosya)
├── requirements.txt
├── baun_logo.png                                   # Logo (kökte)
├── DejaVuSans.ttf                                  # Türkçe destekli font
├── DejaVuSans-Bold.ttf
├── .streamlit/
│   └── config.toml                                 # Açık tema
└── data/
    └── son_3_yil_program_basari_istatistigi.xlsx   # Veri
```

> **Not:** `app.py` dosya yollarını esnek arar. `baun_logo.png` ve fontlar
> `assets/` klasöründe de olabilir, kökte de — kod ikisini de bulur.

## 🚀 Yerel Çalıştırma

```bash
pip install -r requirements.txt
streamlit run app.py
```

## ☁️ Streamlit Cloud Deploy

1. Repo'yu GitHub'a pushla (yukarıdaki tüm dosyalar dahil)
2. [share.streamlit.io](https://share.streamlit.io) → "New app"
3. Main file: `app.py`

**Kritik:** Font dosyaları (`.ttf`) ve logo (`.png`) **MUTLAKA** pushlanmalı.
`git ls-files` ile doğrula:
```bash
git ls-files | grep -E "(ttf|png|xlsx)"
```

## 🎨 Özellikler

- **3 görüntüleme modu:** Tek program, Toplu görünüm, Karşılaştırma
- **PNG indirme:** Her grafiğin altında. Logo + BAUN başlığı + grafik +
  büyük yazılı metrik panel hep bir arada.
- **Türkçe karakter garantisi:** Repo-içi font sayesinde "İ, ş, ğ" sorunsuz.
- **Açık tema zorunlu:** Cloud'un dark mode'una takılı kalmaz.
- **Y ekseni ters:** Sıra düşüşü grafikte YUKARI gider (iyileşme).
