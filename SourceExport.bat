@echo off
cd %~dp0
set baseDir=%~dp0

rd /s /q "%baseDir%Source\01.���ς���H���Z�o�p.xlsm"
rd /s /q "%baseDir%Source\10.�݌v��_��{�y�������V�X�e���z.xlsm"
rd /s /q "%baseDir%Source\11.�݌v��_DB�y�������V�X�e���z.xlsm"
rd /s /q "%baseDir%Source\12.�݌v��_IF�y�������V�X�e���z.xlsm"
rd /s /q "%baseDir%Source\41.�e�X�g���ʕ񍐏��y�������z.xlsm"

cd ..
cscript //nologo vbac.wsf decombine /binary:Template /source:Template/Source
