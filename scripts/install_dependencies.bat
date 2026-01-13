@echo off
REM ============================================
REM Installation des dependances MCP document-operations
REM Version: 0.3.1
REM ============================================

echo.
echo === Installation MCP document-operations v0.3.1 ===
echo.

REM Detecter le chemin du MCP
set "MCP_PATH=%USERPROFILE%\document-edit-mcp"

if not exist "%MCP_PATH%" (
    echo ERREUR: Dossier MCP non trouve: %MCP_PATH%
    echo Veuillez d'abord cloner ou copier le MCP dans ce dossier.
    pause
    exit /b 1
)

echo Chemin MCP: %MCP_PATH%
cd /d "%MCP_PATH%"

REM Verifier si le venv existe
if not exist ".venv" (
    echo Creation de l'environnement virtuel...
    python -m venv .venv
)

echo.
echo Installation des dependances...
echo.

.\.venv\Scripts\pip.exe install --upgrade pip
.\.venv\Scripts\pip.exe install -e .

echo.
echo === Installation terminee ===
echo.
echo Dependances installees:
echo   - python-docx (Word)
echo   - pandas, openpyxl (Excel)
echo   - reportlab (PDF)
echo   - docx2pdf (conversion)
echo   - send2trash (corbeille)
echo.
echo IMPORTANT: Redemarrez Claude Desktop pour appliquer les changements.
echo.
pause
