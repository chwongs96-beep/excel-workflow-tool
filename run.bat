@echo off
echo Installing Excel Workflow Tool dependencies...
pip install -r requirements.txt

echo.
echo Starting Excel Workflow Tool...
python main.py
