@echo off
call "C:\Program Files\Microsoft Visual Studio\2022\Community\VC\Auxiliary\Build\vcvars64.bat">nul
rc /nologo textbox.rc
cl /nologo /GS- /GF /Os /O1 /DNDEBUG /TC show.c /link /nodefaultlib /entry:main textbox.res advapi32.lib comdlg32.lib kernel32.lib user32.lib
pause
