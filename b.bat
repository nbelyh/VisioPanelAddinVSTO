@echo off

%WINDIR%\Microsoft.NET\Framework\v3.5\MSBuild.exe %* /p:Platform=x86 /p:Configuration=Release
%WINDIR%\Microsoft.NET\Framework\v3.5\MSBuild.exe %* /p:Platform=x64 /p:Configuration=Release

rem %WINDIR%\Microsoft.NET\Framework\v3.5\MSBuild.exe %* /p:Platform=x86 /p:Configuration=Debug
rem %WINDIR%\Microsoft.NET\Framework\v3.5\MSBuild.exe %* /p:Platform=x64 /p:Configuration=Debug
