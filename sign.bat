"C:\Program Files (x86)\Windows Kits\10\bin\x64\signtool.exe" ^
sign /n "Nikolay Belykh" /v ^
/fd sha256 ^
/tr http://timestamp.digicert.com /td sha256 ^
/d "Visio TwoPointMove Addin" ^
/du "http://unmanagedvisio.com" ^
"ship\*.msi" 
