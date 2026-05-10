chcp 65001 >nul
title 快速还原Killer - by Doubao AI and mc_lhz
echo 快速还原Killer，通过替换fadisk.sys使快速还原失效
fltmc >nul 2>&1 || (echo Please run as Administrator!  & exit)

echo 1 - Replace fadisk.sys
echo 2 - Restore fadisk.sys
set /p "ch=Choose 1/2: "
set /p "drv=System drive (no colon, e.g. C): "
set "target=%drv%:\Windows\System32\drivers"
set "sys=fadisk.sys"
set "blank=blank.sys"

if "%ch%"=="1" (
echo Replacing...
echo "%target%\%sys%" "%sys%.FastRestoreKiller"
copy "%blank%" "%target%\%sys%" /y
echo Replace done!
)

if "%ch%"=="2" (
echo Restoring...
del /f /q "%target%\%sys%"
ren "%target%\%sys%.FastRestoreKiller"
echo "%target%\%sys%.FastRestoreKiller" "%sys%"
echo Restore done!
)

pause >nul