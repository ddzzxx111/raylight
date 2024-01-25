
@echo off

rem %1 %2 mshta vbscript:CreateObject("Shell.Application").ShellExecute("%~s0","gotore :runas","","runas",0)(window.close)&goto :eof
rem :runas

set "current_dir=%~dp0"
for %%i in ("%current_dir%..") do for %%j in (%%~dpi.) do set "parent_dir=%%~dpnxj"
set "py_path=%parent_dir%\rl_app\Manday\python\python39"

if exist C:\"Program Files"\Raylight\rl_app\ (
set Path=C:\"Program Files"\Raylight\rl_app\Manday\python\python39;
set PYTHONPATH=C:\"Program Files"\Raylight\rl_app\Manday\python\python39\Lib;
set PYTHONPATH=C:\"Program Files"\Raylight\rl_app\Manday\python\python39\Lib\site-packages;

python.exe "C:\Program Files\Raylight\rl_app\showDashboard\bin\showDashboard.py"
) else (
set "current_dir=%~dp0"
rem for %%i in ("%current_dir%..") do for %%j in (%%~dpi.) do set "parent_dir=%%~dpnxj"

set "bat_dir=%~dp0"
set Path=%py_path%
set PYTHONPATH="%py_path%\Lib\site-packages"
set PYTHONPATH="%py_path%\Lib\"
set dashboard_path=%~dp0bin\showDashboard.py

rem python.exe %manday_path%
python.exe %~dp0bin\showDashboard.py
)

python --version
pause