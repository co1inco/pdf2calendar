rd /s /q "build"
py -m setup build
move ../build .

mkdir "build/exe.win-amd64-3.6/tcl/tk8.6"
cp "%TK_LIBRARY%" build/exe.win-amd64-3.6/tcl/tk8.6/ -r

makensis setup.nsi
rm pdf2calendar.exe
rename main.exe pdf2calendar.exe

pause