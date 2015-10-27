@echo off
D:
md  c:\his4tools
cd  c:\his4tools
taskkill /IM his4tools.exe /F 
if exist canshu.ini goto canshu else
copy \\192.189.101.98\backup\his4tools\*.*  /y
copy ms*.* %windir%\system32\ /y
Regsvr32 %windir%\system32\msadodc.ocx /s
Regsvr32 %windir%\system32\mscomm32.ocx /s
Regsvr32 %windir%\system32\msdatgrd.ocx /s
Regsvr32 %windir%\system32\msstdfmt.dll /s
copy ms*.* %windir%\syswow64\ /y
Regsvr32 %windir%\syswow64\msadodc.ocx /s
Regsvr32 %windir%\syswow64\mscomm32.ocx /s
Regsvr32 %windir%\syswow64\msdatgrd.ocx /s
Regsvr32 %windir%\syswow64\msstdfmt.dll /s
Regsvr32 %windir%\syswow64\Richtx32.ocx /s
start his4tools.exe
exit

:canshu
copy canshu.ini canshu.ini.bak /y
copy \\192.189.101.98\backup\his4tools\*.*  /y
copy canshu.ini.bak canshu.ini /y
copy ms*.* %windir%\system32\ /y
Regsvr32 %windir%\system32\msadodc.ocx /s
Regsvr32 %windir%\system32\mscomm32.ocx /s
Regsvr32 %windir%\system32\msdatgrd.ocx /s
Regsvr32 %windir%\system32\msstdfmt.dll /s
copy ms*.* %windir%\syswow64\ /y
Regsvr32 %windir%\syswow64\msadodc.ocx /s
Regsvr32 %windir%\syswow64\mscomm32.ocx /s
Regsvr32 %windir%\syswow64\msdatgrd.ocx /s
Regsvr32 %windir%\syswow64\msstdfmt.dll /s
Regsvr32 %windir%\syswow64\Richtx32.ocx /s
start his4tools.exe
exit