@echo off
cd /d "%~dp0"
set CSC=C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe
if not exist "%CSC%" set CSC=C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe

echo ============================================
echo   Excel DDE Fix v3 - Build
echo ============================================

"%CSC%" /nologo /target:winexe /optimize+ /r:System.dll /r:System.Drawing.dll /r:System.Windows.Forms.dll /out:ExcelDdeFix.exe src\RegEngine.cs src\ExcelDdeGui.cs src\Cli.cs src\HtmlReport.cs

if errorlevel 1 (
  echo BUILD FAILED
  exit /b 1
)

echo BUILD OK - ExcelDdeFix.exe created
echo Source files: RegEngine.cs ExcelDdeGui.cs Cli.cs HtmlReport.cs
