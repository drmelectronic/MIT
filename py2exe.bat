rmdir build /Q /S
rmdir dist /Q /S
python setup.py py2exe
cd dist
MIT03.exe
cd ..
start compilador.iss