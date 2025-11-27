$ErrorActionPreference = "Stop"

$odt = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool.exe"
Invoke-WebRequest $odt -OutFile C:\odt.exe
Start-Process C:\odt.exe -ArgumentList "/quiet /extract:C:\Office" -Wait

@"
<Configuration>
  <Add OfficeClientEdition="64">
    <Product ID="O365ProPlusRetail">
      <Language ID="en-us"/>
    </Product>
  </Add>
  <Display Level="None" AcceptEULA="True"/>
</Configuration>
"@ | Out-File C:\Office\config.xml -Encoding utf8

Start-Process C:\Office\setup.exe -ArgumentList "/configure C:\Office\config.xml" -Wait
