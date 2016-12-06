echo off
cls

%windir%\Microsoft.NET\Framework\v2.0.50727\regasm.exe Ti9LeXML.dll /tlb:Ti9LeXML.tlb  /codebase /unregister
del /q /f Ti9LeXML.tlb


rem %windir%\Microsoft.NET\Framework\v2.0.50727\regasm.exe Ti9LeXML.dll /tlb: Ti9LeXML.tlb /register

rem regasm Ti9LeXML.dll /tlb:Ti9LeXML.tlb

rem regasm Ti9LeXML.dll /tlb:Ti9LeXML.tlb /codebase /register


pause