@echo off
python -m venv excel_env
call excel_env\Scripts\activate.bat
pip install -r requirements.txt
echo Environment ready! Run 'python GenerateZTable.py'
pause