python3 -m venv .venv
source .venv/bin/activate
pip install --upgrade pip
pip install -r requirements.txt
pip install pyinstaller==6.10.0

pyinstaller --clean -y build.spec

echo "Built in: dist/"
