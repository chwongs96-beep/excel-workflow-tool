@echo off
echo Installing dependencies...
pip install -r requirements.txt

echo.
echo Running regression tests (CSV/XLS -> XLSX)...
python -m unittest tests.test_regression -v
