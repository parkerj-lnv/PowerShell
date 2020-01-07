<#
	.SYNOPSIS
		Create Lenovo System Update Application

	.DESCRIPTION
		Script will download the latest version of System Update from Lenovo's support site, create a ConfigMgr Application/Deployment Type, and distribute to a Distribution Point

	.PARAMETER SourcePath
		Source location System Update executable will be downloaded to

	.PARAMETER DistributionPoint
		FQDN Name of a ConfigMgr Distribution Point

    .NOTE
        Internet Explorer will need to be opened before running due to the first run wizard.

#>

[CmdletBinding()]
param
(
    [parameter(Mandatory = $true, HelpMessage = "Specify the UNC path where System Update will be downloaded to")]
    [ValidateNotNullOrEmpty()]
    [string]$SourcePath,
    [parameter(Mandatory = $true, HelpMessage = "Specify FQDN of a Distribution Point")]
    [ValidateNotNullOrEmpty()]
    [string]$DistributionPoint

)

# Parse the TVSU web page and grab the latest download link
$path = "https://support.lenovo.com/downloads/ds012808"
$ie = New-Object -ComObject InternetExplorer.Application
$ie.visible = $false
$ie.navigate($path)
while ($ie.ReadyState -ne 4) { Start-Sleep -Milliseconds 100 }
$document = $ie.document
$exeURL = $document.links | ? { $_.href.EndsWith(".exe") } | % { $_.href }
$exe = $exeURL.Split('/')[5]
$exeVer = $exe.Split('_')[2].TrimEnd('.exe')
# Downloading to source location
Invoke-WebRequest -Uri $exeURL -OutFile "$SourcePath\$exe"
$ie.Quit

# Saving the Thumbprint of the System Update certificate as a variable.
$thumbprint = "0DB6ED63773CD32463B7B3B5A7392F737DF81D10"

# Compare Certificate Thumbprints to verify authenticity.  Script errors out if thumbprints do not match.
If ((Get-AuthenticodeSignature -FilePath $SourcePath\$exe).SignerCertificate.Thumbprint -ne $thumbprint) {
    Write-Error "Certificate thumbprints do not match.  Exiting out" -ErrorAction Stop
}

# Import ConfigMgr PS Module
Import-Module $env:SMS_ADMIN_UI_PATH.Replace("bin\i386", "bin\ConfigurationManager.psd1") -Force

# Connect to ConfigMgr Site
$SiteCode = $(Get-WmiObject -ComputerName "$ENV:COMPUTERNAME" -Namespace "root\SMS" -Class "SMS_ProviderLocation").SiteCode
if (!(Get-PSDrive $SiteCode)) { }
New-PSDrive -Name $SiteCode -PSProvider "AdminUI.PS.Provider\CMSite" -Root "$ENV:COMPUTERNAME" -Description "Primary Site Server" -ErrorAction SilentlyContinue
Set-Location "$SiteCode`:"

# Create the System Update App
$app = New-CMApplication -Name "System Update" `
    -Publisher "Lenovo" `
    -SoftwareVersion "$exeVer" `
    -LocalizedName "Lenovo System Update" `
    -LocalizedDescription "System Update enables IT administrators to distribute updates for software, drivers, and BIOS in a managed environment from a local server." `
    -LinkText "https://support.lenovo.com/downloads/ds012808" `
    -Verbose

# Create Registry detection clause
$clause = New-CMDetectionClauseRegistryKeyValue -ExpressionOperator IsEquals `
    -Hive LocalMachine `
    -KeyName "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\TVSU_is1" `
    -PropertyType Version `
    -ValueName "DisplayVersion" `
    -Value:$true `
    -ExpectedValue "$exeVer" `
    -Verbose

# Add Deployment Type
$app | Add-CMScriptDeploymentType -DeploymentTypeName "System Update-$exeVer" `
    -ContentLocation $SourcePath `
    -InstallCommand "$exe /verysilent /norestart" `
    -UninstallCommand "unins000.exe /verysilent /norestart" `
    -UninstallWorkingDirectory "%PROGRAMFILES(X86)%\Lenovo\System Update" `
    -AddDetectionClause $clause `
    -InstallationBehaviorType InstallForSystem `
    -Verbose

# Distribute app to Distribution Point
$app | Start-CMContentDistribution -DistributionPointName $DistributionPoint -ErrorAction SilentlyContinue -Verbose

Set-Location -Path $env:HOMEDRIVE