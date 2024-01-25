
@echo off

rem %1 %2 mshta vbscript:CreateObject("Shell.Application").ShellExecute("%~s0","gotore :runas","","runas",0)(window.close)&goto :eof
rem :runas


rem echo %PYTHONPATH%
rem echo %manday_path%
rem echo %Path%

if exist C:\"Program Files"\Raylight\rl_app\ (
set Path=C:\"Program Files"\Raylight\rl_app\Manday\python\python39;
set PYTHONPATH=C:\"Program Files"\Raylight\rl_app\Manday\python\python39\Lib;
set PYTHONPATH=C:\"Program Files"\Raylight\rl_app\Manday\python\python39\Lib\site-packages;

python.exe "C:\Program Files\Raylight\rl_app\Manday\bin\manday.py"
) else (
set "current_dir=%~dp0"
rem for %%i in ("%current_dir%..") do for %%j in (%%~dpi.) do set "parent_dir=%%~dpnxj"

set "bat_dir=%~dp0"
set Path=%~dp0\python\python39
set PYTHONPATH="%bat_dir%\python\python39\Lib\site-packages"
set PYTHONPATH="%bat_dir%\python\python39\Lib\"
set manday_path=%~dp0bin\manday.py

rem python.exe %manday_path%
python.exe %~dp0bin\manday.py
)

python --version
pause