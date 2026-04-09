@echo off
setlocal
set EXE=netmon.exe
set SRC=src\main.c

echo Building %EXE% ...
gcc -std=c11 -Os -s -mwindows -o %EXE% %SRC% -lws2_32 -liphlpapi -lmpr -lshell32
if errorlevel 1 (
    echo Build failed.
    exit /b 1
)

echo Done: %EXE%
endlocal
