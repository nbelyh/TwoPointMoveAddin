@echo off

"C:\Program Files (x86)\MSBuild\14.0\Bin\MSBuild.exe" %* /p:Platform=x86 /p:Configuration=Release
"C:\Program Files (x86)\MSBuild\14.0\Bin\MSBuild.exe" %* /p:Platform=x64 /p:Configuration=Release
