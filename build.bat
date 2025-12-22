@echo off
echo Installing build dependencies...
pip install pyinstaller

echo.
echo Building Excel Workflow Tool...
echo This may take a few minutes.

if exist dist rmdir /s /q dist
if exist build rmdir /s /q build

pyinstaller main.py ^
    --name "ExcelWorkflowTool" ^
    --windowed ^
    --onefile ^
    --add-data "assets;assets" ^
    --add-data "templates;templates" ^
    --clean

echo.
if exist "dist\ExcelWorkflowTool.exe" (
    echo Build successful!
    echo The executable is located at: dist\ExcelWorkflowTool.exe
) else (
    echo Build failed. Please check the error messages above.
)
pause
