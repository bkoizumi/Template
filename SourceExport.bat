@echo off
cd %~dp0
set baseDir=%~dp0

rd /s /q "%baseDir%Source\01.見積もり工数算出用.xlsm"
rd /s /q "%baseDir%Source\10.設計書_基本【＊＊＊システム】.xlsm"
rd /s /q "%baseDir%Source\11.設計書_DB【＊＊＊システム】.xlsm"
rd /s /q "%baseDir%Source\12.設計書_IF【＊＊＊システム】.xlsm"
rd /s /q "%baseDir%Source\41.テスト結果報告書【＊＊＊】.xlsm"

cd ..
cscript //nologo vbac.wsf decombine /binary:Template /source:Template/Source
