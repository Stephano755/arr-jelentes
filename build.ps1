# ARR Jelentés – EXE építése (Windows PowerShell)
# Előfeltétel: Python 3.12, és a projekt mappájában "sablon.xlsm" néven az Excel sablon.
if (-not (Test-Path .\sablon.xlsm)) { Write-Error "Hiányzik a sablon.xlsm (az Excel sablon) a projekt mappájából."; exit 1 }
if (-not (Test-Path .\.venv)) { python -m venv .venv }
.\.venv\Scripts\python.exe -m pip install -q --disable-pip-version-check -r requirements.txt
.\.venv\Scripts\pyinstaller.exe --noconfirm --clean --onefile --noconsole --name ARR_Jelentes --add-data "sablon.xlsm;." arr_jelentes.py
Write-Host "Kész: dist\ARR_Jelentes.exe"
