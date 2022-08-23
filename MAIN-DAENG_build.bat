set root=C:\Users\enmso\anaconda3
call %root%\Scripts\activate.bat %root%

call conda activate youtube-to-mp3-downloader
call pyinstaller -n "ENM-YTMD" --onefile --noconsole main.py
call move "dist\ENM-YTMD.exe"