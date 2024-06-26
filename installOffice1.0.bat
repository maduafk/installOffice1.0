@echo off

echo "Automacao feita por @maduafk"
mkdir "C:\MS Office Setup"

REM download e extração do office deployments
powershell -Command "(New-Object System.Net.WebClient).DownloadFile('https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_17531-20046.exe', 'C:\MS Office Setup\officedeploymenttool_17531-20046.exe')"

cd "C:\MS Office Setup"
officedeploymenttool_17531-20046.exe /extract:"C:\MS Office Setup" /quiet

REM config.xml
(
    echo ^<Configuration ID="3d44ecdf-3c94-466d-b7a9-98659f04bf8b"^>
    echo    ^<Add OfficeClientEdition="64" Channel="PerpetualVL2021"^>
    echo        ^<Product ID="ProPlus2021Volume" PIDKEY="FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH"^>
    echo            ^<Language ID="MatchOS"/^>
    echo            ^<ExcludeApp ID="Lync"/^>
    echo        ^</Product^>
    echo    ^</Add^>
    echo    ^<Property Name="SharedComputerLicensing" Value="0"/^>
    echo    ^<Property Name="FORCEAPPSHUTDOWN" Value="FALSE"/^>
    echo    ^<Property Name="DeviceBasedLicensing" Value="0"/^>
    echo    ^<Property Name="SCLCacheOverride" Value="0"/^>
    echo    ^<Property Name="AUTOACTIVATE" Value="1"/^>
    echo    ^<Updates Enabled="TRUE"/^>
    echo    ^<RemoveMSI/^>
    echo    ^<AppSettings^>
    echo        ^<User Key="software\microsoft\office\16.0\excel\options" Name="defaultformat" Value="51" Type="REG_DWORD" App="excel16" Id="L_SaveExcelfilesas"/^>
    echo        ^<User Key="software\microsoft\office\16.0\powerpoint\options" Name="defaultformat" Value="27" Type="REG_DWORD" App="ppt16" Id="L_SavePowerPointfilesas"/^>
    echo        ^<User Key="software\microsoft\office\16.0\word\options" Name="defaultformat" Value="" Type="REG_SZ" App="word16" Id="L_SaveWordfilesas"/^>
    echo    ^</AppSettings^>
    echo ^</Configuration^>
) > "C:\MS Office Setup\config.xml"

cd "C:\MS Office Setup"
setup.exe /configure config.xml

REM 'configurar' o office apos a instalação
cd /d %ProgramFiles%\Microsoft Office\Office16
cd /d %ProgramFiles(x86)%\Microsoft Office\Office16
for /f %%x in ('dir /b ..\root\Licenses16\proplusvl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x"
cscript ospp.vbs /inpkey:XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99
cscript ospp.vbs /unpkey:BTDRB >nul
cscript ospp.vbs /unpkey:KHGM9 >nul
cscript ospp.vbs /unpkey:CPQVG >nul
cscript ospp.vbs /sethst:kms8.msguides.com
cscript ospp.vbs /setprt:1688
cscript ospp.vbs /act

exit
