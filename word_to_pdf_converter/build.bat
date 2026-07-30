@echo off
REM Word'den PDF'e Donusturucu - Tek dosya (.exe) derleme betigi.
REM Bu betigi Windows uzerinde, bu klasorden calistirin.

python -m pip install --upgrade pip
python -m pip install -r requirements.txt pyinstaller

pyinstaller --onefile --windowed --name WordToPDFDonusturucu word_to_pdf_gui.py

echo.
echo Derleme tamamlandi. Tek dosyalik uygulama: dist\WordToPDFDonusturucu.exe
pause
