python -m venv .venv
.venv\Scripts\Activate.ps1
pip install --upgrade pip
pip install -r requirements.txt
pip install pyinstaller==6.10.0

pyinstaller --clean -y build.spec

Write-Host "EXE built at: dist\Mplus2Excel_Structural.exe"
