@echo off
call "C:\Program Files (x86)\Microsoft Visual Studio\2019\Community\VC\Auxiliary\Build\vcvars32.bat">nul
rc /nologo textbox.rc
cl /nologo /GS- /GF /Os /O1 /DNDEBUG /DTXT_NOBEEPS /Tp show.c textbox.res advapi32.lib comdlg32.lib kernel32.lib user32.lib /link /nodefaultlib /entry:main
pause
