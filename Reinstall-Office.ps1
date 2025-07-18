# Define working directory and paths
$workingDir = "$env:TEMP\ODT"
$setupPath = "$workingDir\setup.exe"
$xmlPathUninstall = "$workingDir\uninstall.xml"
$xmlPathInstall = "$workingDir\install.xml"

# Create working directory
New-Item -ItemType Directory -Force -Path $workingDir | Out-Null

# Download Office Deployment Tool setup.exe directly
Invoke-WebRequest -Uri "https://officecdn.microsoft.com/pr/wsus/setup.exe" -OutFile $setupPath

# Ensure setup.exe was downloaded
if (-Not (Test-Path $setupPath)) {
    Write-Output "ERROR: setup.exe not found at $setupPath"
    Exit 1
}

# Create uninstall config XML
@"
<Configuration>
  <Remove All="TRUE" />
  <Display Level="None" AcceptEULA="TRUE" />
</Configuration>
"@ | Out-File -FilePath $xmlPathUninstall -Encoding UTF8

# Run Office uninstall
Start-Process -FilePath $setupPath -ArgumentList "/configure $xmlPathUninstall" -Wait

# Wait before reinstall
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

# Optional cleanup
Remove-Item -Path $workingDir -Recurse -Force -ErrorAction SilentlyContinue
