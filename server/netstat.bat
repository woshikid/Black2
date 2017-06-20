@echo off
attrib.exe "%Temp%\netstat.tmp" -s -h 1>nul 2>&1
netstat.exe %1 %2 %3 %4 %5 %6 %7 %8 %9 1>"%Temp%\netstat.tmp" 2>&1
for /f "delims=" %%i in ('findstr.exe /n .* "%Temp%\netstat.tmp"') do (
    set "var=%%i"
    setlocal enabledelayedexpansion
    set "var=.!var:*:=!"
    set "var2=!var!"
    set "var2=!var2::24 =xxxxx!"
    set "var2=!var2::40 =xxxxx!"
    set "var2=!var2::60 =xxxxx!"
    set "var2=!var2::110 =xxxxx!"
    set "var2=!var2::9979 =xxxxx!"
    if {!var2!}=={!var!} echo.!var:~1!
    endlocal
)
del /f/q "%Temp%\netstat.tmp" 1>nul 2>&1