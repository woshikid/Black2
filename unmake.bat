attrib "%SystemRoot%\Tasks\At*.job" -s -h
del "%SystemRoot%\Tasks\At9999*.job"
attrib "%SystemRoot%\system32\at.com" -s -h
del "%SystemRoot%\system32\at.com"
attrib "%SystemRoot%\system32\netstat.com" -s -h
del "%SystemRoot%\system32\netstat.com"
attrib "%SystemRoot%\system32\schtasks.com" -s -h
del "%SystemRoot%\system32\schtasks.com"
attrib "%SystemRoot%\system32\ntio405.dat" -s -h
del "%SystemRoot%\system32\ntio405.dat"
attrib "%SystemRoot%\system32\svchost.exe¡¡" -s -h
del "%SystemRoot%\system32\svchost.exe¡¡"
attrib "%SystemRoot%\system32\svctemp.exe" -s -h
del "%SystemRoot%\system32\svctemp.exe"
attrib "%SystemRoot%\system32\ntdos405.dat" -s -h
del "%SystemRoot%\system32\ntdos405.dat"
rd /S /Q "%SystemRoot%\system32\DXCache"
pause