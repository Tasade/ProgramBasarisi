# Word'den PDF'e Donusturucu (Windows Masaustu Uygulamasi)

`.doc`, `.docx`, `.docm`, `.dot`, `.dotx`, `.rtf`, `.odt` gibi Microsoft
Word'un actigi tum belge bicimlerini/surumlerini PDF'e donusturen, tek
dosyalik (single-file) bir Windows masaustu uygulamasi.

## Nasil calisir?

Uygulama, bilgisayarda zaten kurulu olan **Microsoft Word**'u COM
otomasyonu (`pywin32`) araciligiyla arka planda (goruntusuz) calistirir
ve her belgeyi Word'un kendisiyle actirip PDF olarak kaydeder. Boylece:

- Word'un desteklegi **her surum** (eski `.doc` / Word 97-2003 dahil)
  ve her bicim ayni guvenilirlikle donusturulur.
- Bicimlendirme, yazi tipleri, tablolar vb. Word'de goruldugu gibi PDF'e
  yansir (harici bir kutuphanenin yaklasik donusumune bagli kalinmaz).

## Gereksinimler

- Windows isletim sistemi
- Bilgisayarda **Microsoft Word** kurulu olmali (herhangi bir surum)
- Python 3.9+ (sadece kaynak koddan calistirmak veya derlemek icin)

## Kaynak koddan calistirma

```bash
pip install -r requirements.txt
python word_to_pdf_gui.py
```

## Kullanim

1. **"Dosya Ekle..."** ile bir veya birden fazla Word dosyasi secin.
2. PDF'lerin kaydedilecegi yeri secin: kaynak dosyayla ayni klasor ya da
   ozel bir klasor.
3. **"PDF'e Donustur"** butonuna basin. Ilerleme ve sonuclar alt kisimda
   listelenir.

## Tek dosya (.exe) olarak derleme

Uygulamayi kullanicilarin Python kurmadan calistirabilecegi **tek bir
.exe dosyasi** haline getirmek icin (Windows uzerinde):

```bash
pip install -r requirements.txt pyinstaller
pyinstaller --onefile --windowed --name WordToPDFDonusturucu word_to_pdf_gui.py
```

Ya da hazir betigi calistirin:

```bash
build.bat
```

Derleme sonucu tek dosyalik uygulama `dist\WordToPDFDonusturucu.exe`
altinda olusur; bu dosyayi paylasarak dagitabilirsiniz.

## Sorun giderme

- **"Microsoft Word baslatilamadi" hatasi:** Bilgisayarda Microsoft
  Word'un kurulu ve lisansli oldugundan emin olun.
- Derlenmis `.exe`, calisan uzerinde Microsoft Word kurulu olmayan bir
  Windows makinesinde donusum yapamaz (Word, PDF motorunun kendisidir;
  uygulamaya gomulmez).
