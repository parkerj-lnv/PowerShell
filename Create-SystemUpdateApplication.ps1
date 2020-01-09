<#
	.SYNOPSIS
		Create Lenovo System Update Application

	.DESCRIPTION
		Script will download the latest version of System Update and Thin Installer (if specified) from Lenovo's support site, create a ConfigMgr Application/Deployment Type, and distribute to a Distribution Point

	.PARAMETER SystemUpdateSourcePath
        Source location System Update executable will be downloaded to

    .SWITCH ThinInstaller
        Downloads the latest version of Thin Installer

    .PARAMETER ThinInstallerSourcePath
		Source location Thin Installer executable will be downloaded to

	.PARAMETER DistributionPoint
		FQDN Name of a ConfigMgr Distribution Point

    .NOTE
    Run script as Administrator on Site Server
    Turn off Internet Explorer Enhanced Security Control for Administrators prior to running

    .EXAMPLE
    .\Create-SystemUpdateApplication -SystemUpdateSourcePath "\\Share\Software\Lenovo\SystemUpdate\5.07.88" -DistributionPoint "\\dp.local"

    .EXAMPLE
    .\Create-SystemUpdateApplication -SystemUpdateSourcePath "\\Share\Software\Lenovo\SystemUpdate\5.07.88" -ThinInstaller -ThinInstallerSourcePath ""\\Share\Software\Lenovo\ThinInstaller\1.3.00018" -DistributionPoint "\\dp.local"

#>

[CmdletBinding()]
param
(
    [Parameter(Mandatory = $true, HelpMessage = "Specify the UNC path where System Update will be downloaded to")]
    [ValidateNotNullOrEmpty()]
    [String]$SystemUpdateSourcePath,

    [Parameter(ParameterSetName = 'ThinInstaller Set 1')]
    [switch]$ThinInstaller,

    [Parameter(ParameterSetName = 'ThinInstaller Set 1', Mandatory, HelpMessage = "Specify the UNC path where ThinInstaller will be downloaded to")]
    [Parameter(ParameterSetName = 'ThinInstaller Set 2')]
    [String]$ThinInstallerSourcePath,

    [Parameter(Mandatory = $true, HelpMessage = "Specify FQDN of a Distribution Point")]
    [ValidateNotNullOrEmpty()]
    [String]$DistributionPoint

)

# Parse the TVSU web page and grab the latest download link
$path = "https://support.lenovo.com/solutions/ht037099"
$ie = New-Object -ComObject InternetExplorer.Application
$ie.visible = $false
$ie.navigate($path)
while ($ie.ReadyState -ne 4) { Start-Sleep -Milliseconds 100 }
$document = $ie.document
$suExeURL = $document.links | ? { $_.href.Contains("system_update") -and $_.href.EndsWith(".exe") } | % { $_.href }
$suExe = $suExeURL.Split('/')[5]
$suExeVer = $suExe.Split('_')[2].TrimEnd('.exe')
# Downloading System Update to source location
Invoke-WebRequest -Uri $suExeURL -OutFile "$SystemUpdateSourcePath\$suExe"
$ie.Quit

# Saving the Thumbprint of the System Update certificate as a variable.
$suThumbprint = "0DB6ED63773CD32463B7B3B5A7392F737DF81D10"

If ($ThinInstaller) {
    $tiExeURL = $document.links | ? { $_.href.Contains("thininstaller") -and $_.href.EndsWith(".exe") } | % { $_.href }
    $tiExe = $tiExeURL.Split('/')[5]
    # Downloading System Update to source location
    Invoke-WebRequest -Uri $tiExeURL -OutFile "$ThinInstallerSourcePath\$tiExe"
    $tiExeVerRaw = (Get-ChildItem -Path "$ThinInstallerSourcePath\$tiExe").VersionInfo.FileVersionRaw
    $tiExeVer = "$($tiExeVerRaw.Major).$($tiExeVerRaw.Minor).$($tiExeVerRaw.Build)"
    $ie.Quit

    # Saving the Thumbprint of the ThinInstaller certificate as a variable.
    $tiThumbprint = "CC5EE80524D43ACD5A32AB1F3A9D163CEE924443"

}

# Compare Certificate Thumbprints to verify authenticity.  Script errors out if thumbprints do not match.
If ((Get-AuthenticodeSignature -FilePath $SystemUpdateSourcePath\$suExe).SignerCertificate.Thumbprint -ne $suThumbprint) {
    Write-Error "Certificate thumbprints do not match.  Exiting out" -ErrorAction Stop
    If ($ThinInstaller) {
        ((Get-AuthenticodeSignature -FilePath $ThinInstallerSourcePath\$tiExe).SignerCertificate.Thumbprint -ne $tiThumbprint)
        {
            Write-Error "Certificate thumbprints do not match. Exiting out" -ErrorAction Stop
        }

    }
}

# Import ConfigMgr PS Module
Import-Module $env:SMS_ADMIN_UI_PATH.Replace("bin\i386", "bin\ConfigurationManager.psd1") -Force

# Connect to ConfigMgr Site
$SiteCode = $(Get-WmiObject -ComputerName "$ENV:COMPUTERNAME" -Namespace "root\SMS" -Class "SMS_ProviderLocation").SiteCode
If (!(Get-PSDrive $SiteCode)) { }
New-PSDrive -Name $SiteCode -PSProvider "AdminUI.PS.Provider\CMSite" -Root "$ENV:COMPUTERNAME" -Description "Primary Site Server" -ErrorAction SilentlyContinue
Set-Location "$SiteCode`:"

# Create the System Update App
$suApp = New-CMApplication -Name "System Update" `
    -Publisher "Lenovo" `
    -SoftwareVersion "$suExeVer" `
    -LocalizedName "Lenovo System Update" `
    -LocalizedDescription "System Update enables IT administrators to distribute updates for software, drivers, and BIOS in a managed environment from a local server." `
    -LinkText "https://support.lenovo.com/downloads/ds012808" `
    -Verbose

# Create Registry detection clause
$clause1 = New-CMDetectionClauseRegistryKeyValue -ExpressionOperator IsEquals `
    -Hive LocalMachine `
    -KeyName "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\TVSU_is1" `
    -PropertyType Version `
    -ValueName "DisplayVersion" `
    -Value:$true `
    -ExpectedValue "$suExeVer" `
    -Verbose

# Add Deployment Type
$suApp | Add-CMScriptDeploymentType -DeploymentTypeName "System Update-$suExeVer" `
    -ContentLocation $SystemUpdateSourcePath `
    -InstallCommand "$suExe /verysilent /norestart" `
    -UninstallCommand "unins000.exe /verysilent /norestart" `
    -UninstallWorkingDirectory "%PROGRAMFILES(X86)%\Lenovo\System Update" `
    -AddDetectionClause $clause1 `
    -InstallationBehaviorType InstallForSystem `
    -Verbose

If ($ThinInstaller) {

    # Create the ThinInstaller App
    $tiApp = New-CMApplication -Name "Thin Installer" `
        -Publisher "Lenovo" `
        -SoftwareVersion "$tiExeVer" `
        -LocalizedName "Lenovo Thin Installer" `
        -LocalizedDescription "Thin Installer is a smaller version of System Update." `
        -LinkText "https://support.lenovo.com/solutions/ht037099#ti" `
        -Verbose

    # Create Registry detection clause
    $clause2 = New-CMDetectionClauseFile -Path "%PROGRAMFILES(x86)%\Lenovo\ThinInstaller" `
        -FileName "ThinInstaller.exe" `
        -PropertyType Version `
        -Value:$true `
        -ExpressionOperator IsEquals `
        -ExpectedValue $tiExeVer `
        -Verbose

    # Add Deployment Type
    $tiApp | Add-CMScriptDeploymentType -DeploymentTypeName "ThinInstaller-$tiExeVer" `
        -ContentLocation $ThinInstallerSourcePath `
        -InstallCommand "$tiExe /VERYSILENT /SUPPRESSMSGBOXES /NORESTART" `
        -UninstallCommand 'powershell.exe -Command Remove-Item -Path "${env:ProgramFiles(x86)}\Lenovo\ThinInstaller" -Recurse' `
        -AddDetectionClause $clause2 `
        -InstallationBehaviorType InstallForSystem `
        -Verbose
}


# Distribute app to Distribution Point
If (!($ThinInstaller)) {
    $suApp | Start-CMContentDistribution -DistributionPointName $DistributionPoint -ErrorAction SilentlyContinue -Verbose
} If ($ThinInstaller) {
    ($suApp, $tiApp) | Start-CMContentDistribution -DistributionPointName $DistributionPoint -ErrorAction SilentlyContinue -Verbose
}

Set-Location -Path $env:HOMEDRIVE