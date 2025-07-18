# Set working directory
$workingDir = "$env:TEMP\ODT"
$odtUrl = "https://download.microsoft.com/download/2/F/9/2F9FC087-3194-4F29-92FE-7F5F0F6F9F0A/officedeploymenttool_16530-20154.exe"
$setupPath = "$workingDir\setup.exe"
$xmlPathUninstall = "$workingDir\uninstall.xml"
$xmlPathInstall = "$workingDir\install.xml"

# Create working directory
New-Item -ItemType Directory -Force -Path $workingDir | Out-Null

# Download Office Deployment Tool
Invoke-WebRequest -Uri $odtUrl -OutFile "$workingDir\odt.exe"

# Extract ODT
Start-Process -FilePath "$workingDir\odt.exe" -ArgumentList "/quiet /extract:$workingDir" -Wait

# Create uninstall config XML
@"
<Configuration>
  <Remove All="TRUE" />
  <Display Level="None" AcceptEULA="TRUE" />
</Configuration>
"@ | Out-File -FilePath $xmlPathUninstall -Encoding UTF8

# Run Office uninstall
Start-Process -FilePath $setupPath -ArgumentList "/configure $xmlPathUninstall" -Wait

# OPTIONAL: Pause to allow cleanup before reinstall
Start-Sleep -Seconds 20

# Create install config XML
@"
<Configuration>
  <Add OfficeClientEdition="64" Channel="MonthlyEnterprise">
    <Product ID="O365BusinessRetail">
      <Language ID="en-us" />
    </Product>
  </Add>
  <Display Level="None" AcceptEULA="TRUE" />
  <Property Name="AUTOACTIVATE" Value="1" />
</Configuration>
"@ | Out-File -FilePath $xmlPathInstall -Encoding UTF8

# Run Office install
Start-Process -FilePath $setupPath -ArgumentList "/configure $xmlPathInstall" -Wait

# OPTIONAL: Cleanup temp files
Remove-Item -Path $workingDir -Recurse -Force
