@echo off
set /p Date=Enter Report Date (Example 03062021)
set short="C:\YourPath\"
copy /Y %short%SMNEW$\SL%Date%.xlsx %short%\HOLDING\SL1.xlsx
copy /Y %short%SMCREF\CF%Date%.xlsx %short%\HOLDING\CF1.xlsx
copy /Y %short%SMREC\RC%Date%.xlsx %short%\HOLDING\RC1.xlsx
copy /Y %short%SMDEL\DL%Date%.xlsx %short%\HOLDING\DL1.xlsx
copy /Y %short%SMCBEN\CBEN%Date%.xlsx %short%\HOLDING\CBEN1.xlsx

echo %Date% is the report date.>%short%\HOLDING\report_date.txt