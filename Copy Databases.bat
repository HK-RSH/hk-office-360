@echo off
title Copy Databases to GitHub Repo
echo Copying Excel databases...

xcopy "N:\HK-HKG-CBREResearch\Hong Kong\02. Databases\Office Database\Hong Kong Office Database 3.0.xlsm" "C:\Users\ZTang4\OneDrive - CBRE, Inc\Claude\Office\Office Dashboard\hk-office-360\" /Y

xcopy "C:\Users\ZTang4\OneDrive - CBRE, Inc\Claude\Office\Office Databases\Grade A Office Stacking Plan.xlsx" "C:\Users\ZTang4\OneDrive - CBRE, Inc\Claude\Office\Office Dashboard\hk-office-360\" /Y

echo.
echo Done! Now push via GitHub Desktop.
pause
