#Copyright 2019 Quest Software Inc.  ALL RIGHTS RESERVED.

$ErrorActionPreference = "Stop"

$InstallerFolder = "C:\Installer"
$AMPrefix = "ArchiveManager"

$OutputFolder = "C:\Output"


[hashtable]$CongigMainContent = [ordered]@{
    O2013 = "<Configuration Product=`"ProPlusr`">
        {0}
        {1}
        <Setting Id=`"SETUP_REBOOT`" Value=`"IfNeeded`" />
        <Setting Id=`"REBOOT`" Value=`"ReallySuppress`"/>
        <Setting Id=`"AUTO_ACTIVATE`" Value=`"0`" />
        <COMPANYNAME Value=`" MYCOMPANY`" />
        <OptionState Id=`"ACCESSFiles`" State=`"Absent`" Children=`"force`" />
        <OptionState Id=`"EXCELFiles`" State=`"Absent`" Children=`"force`" />
        <OptionState Id=`"XDOCSFiles`" State=`"Absent`" Children=`"force`" />
        <OptionState Id=`"LyncCoreFiles`" State=`"Absent`" Children=`"force`" />
        <OptionState Id=`"OneNoteFiles`" State=`"Absent`" Children=`"force`" />
        <OptionState Id=`"OUTLOOKFiles`" State=`"Local`" Children=`"force`" />
        <OptionState Id=`"PPTFiles`" State=`"Absent`" Children=`"force`" />
        <OptionState Id=`"PubPrimary`" State=`"Absent`" Children=`"force`" />
        <OptionState Id=`"GrooveFiles2`" State=`"Absent`" Children=`"force`" />
        <OptionState Id=`"VisioPreviewerFiles`" State=`"Absent`" Children=`"force`" />
        <OptionState Id=`"WORDFiles`" State=`"Absent`" Children=`"force`" />
        <OptionState Id=`"SHAREDFiles`" State=`"Local`" Children=`"force`" />
        <OptionState Id=`"TOOLSFiles`" State=`"Local`" Children=`"force`" />
    </Configuration>";
    O2019 = "<Configuration>
    {0}
    {1}
    <Add SourcePath=`"{2}`" OfficeClientEdition=`"32`" Channel=`"PerpetualVL2019`">
        <Product ID=`"ProPlus2019Volume`" >
           <Language ID=`"en-us`" />
        </Product>
    </Add>
    <RemoveMSI />
  </Configuration>"
}

[hashtable]$CongigShowUITag = [ordered]@{
    O2013 = "<Display Level=`"none`" CompletionNotice=`"no`" SuppressModal=`"no`" AcceptEula=`"yes`" />";
    O2019 = "<Display Level=`"None`" AcceptEULA=`"True`" />"
}

[hashtable]$CongigLogTag = [ordered]@{
    O2013 = "<Logging Type=`"standard`" Path=`"{0}`" Template=`"Offcie2013Setup.log`" />";
    O2019 = "<Logging Level=`"Standard`" Path=`"{0}`" />"
}

function Get-AMInstallationPath
{
    if([string]::IsNullOrEmpty($script:AMInstallationPath)){
        If (Test-Path -Path Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Quest\ArchiveManager\Installer) {
            $item = Get-Item -Path Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Quest\ArchiveManager\Installer;
        
            $script:AMInstallationPath = $item.GetValue('InstallDirectory');
        }
        ElseIf (Test-Path -Path Registry::HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Quest\ArchiveManager\Installer) {
            $item = Get-Item -Path Registry::HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Quest\ArchiveManager\Installer;
        
            $script:AMInstallationPath = $item.GetValue('InstallDirectory');
        }
    }

    $script:AMInstallationPath
}

function Start-Command {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]
        $Path,
        [Parameter()]
        [string]
        $Arguments,
        [Parameter()]
        [string]
        $WorkingDirectory,
        [Parameter()]
        [switch]
        $UseShell
    )
    $stdout = ""
    $stderr = ""

    Write-Host "Starting process '$Path' with arguments '$Arguments'"
    $pinfo = New-Object System.Diagnostics.ProcessStartInfo
    $pinfo.FileName = $Path
    
    if ($UseShell) {
        $pinfo.RedirectStandardError = $false
        $pinfo.RedirectStandardOutput = $false
        $pinfo.UseShellExecute = $true
    }else {
        $pinfo.RedirectStandardError = $true
        $pinfo.RedirectStandardOutput = $true
        $pinfo.UseShellExecute = $false
    }
    
    $pinfo.Arguments = $Arguments
    if ($WorkingDirectory) {
        $pinfo.WorkingDirectory = $WorkingDirectory
    }
    $p = New-Object System.Diagnostics.Process
    $p.StartInfo = $pinfo
    $p.Start() | Out-Null
    $p.WaitForExit()

    if (-not $UseShell) {
        $stdout = $p.StandardOutput.ReadToEnd()
        $stderr = $p.StandardError.ReadToEnd()
    }

    [pscustomobject]@{
        stdout   = $stdout
        stderr   = $stderr
        ExitCode = $p.ExitCode
    }
}

function New-ConfigFile {
    [CmdletBinding()]
    param (
        [bool]
        $ShowUI,
        [string]
        $OutputPath,
        [ValidateSet("O2013", "O2019")]
        [string]
        $Version,
        [string]
        $InstallPath
    )
    New-Item -ItemType Directory -Force -Path $OutputPath | Out-Null

    $logFolder = Get-Date -Format "yyyyMMddHHmmss"
    $logPath = "$OutputPath\$logFolder"
    New-Item -ItemType Directory -Force -Path $logPath | Out-Null

    $configPath = "$OutputPath\config_$Version.xml"
    $ui = (& { If ($ShowUI) { "" } Else { $CongigShowUITag[$Version] } })
    $log = (& { If ([string]::IsNullOrEmpty($LogPath)) { "" } Else { $CongigLogTag[$Version] -f $LogPath } })

    $config = $CongigMainContent[$Version] -f $ui, $log, $InstallPath
    $config | Out-File -FilePath $configPath

    [pscustomobject]@{
        LogPath    = $logPath
        ConfigPath = $configPath
    }
}

function Get-InstallPackage {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]
        $StoragePath,
        [Parameter(Mandatory)]
        [string]
        $DestinationFolder,
        [Parameter(Mandatory)]
        [object]
        $StorageContext
    )
    
    New-Item -ItemType Directory -Force -Path $DestinationFolder | Out-Null
    if ($DestinationFolder.LastIndexOf('\') -ne $DestinationFolder.Length - 1) {
        $DestinationFolder = "$DestinationFolder\"
    }
    
    Write-Host "Download files from `"$StoragePath`" to `"$DestinationFolder`""
    Get-AzStorageBlob -Container "installer" -Prefix $StoragePath -Context $StorageContext | Get-AzStorageBlobContent -Destination $DestinationFolder -Force
}


function Install-Office2013 {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]
        $InstallerFolder
    )

    $o2013OutputFolder = "$OutputFolder\2013"
    $paths = New-ConfigFile -OutputPath $o2013OutputFolder -ShowUI $false -Version "O2013"
    Write-Host "Paths: $paths"

    Write-Host "Starting install Office 2013"
    $result = Start-Command -Path "$InstallerFolder\setup.exe" -Arguments "/config `"$($paths.ConfigPath)`""

    [pscustomobject]@{
        LogPath  = $paths.LogPath
        ExitCode = $result.ExitCode
    }
}

function Install-Office2013SP1 {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]
        $InstallerFolder,
        [Parameter(Mandatory)]
        [string]
        $KBFolder
    )

    $o2013OutputFolder = "$OutputFolder\2013SP1"
    $paths = New-ConfigFile -OutputPath $o2013OutputFolder -ShowUI $false -Version "O2013"
    Write-Host "Paths: $paths"

    Write-Host "Starting install Office 2013 SP1"
    $result = Start-Command -Path "$InstallerFolder\setup.exe" -Arguments "/config `"$($paths.ConfigPath)`""

    if ($result.ExitCode -eq 0) {
        Write-Host "Install KBs..."
        $kbs = Get-ChildItem -Path $KBFolder -Filter "*.exe" | Sort-Object -Property Name
        $kbs | %{
            $result = Start-Command -Path $_.FullName -Arguments "/q /z /log:`"$($paths.LogPath)\$($_.Name).log`""
            if ($result.ExitCode -ne 0) {
                Write-Host "Failed to install KB $($_.Name)"
                break
            }
        }
    }

    [pscustomobject]@{
        LogPath  = $paths.LogPath
        ExitCode = $result.ExitCode
    }
}

function Install-Office2019 {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]
        $InstallerFolder
    )

    $o2019OutputFolder = "$OutputFolder\2019"
    $paths = New-ConfigFile -OutputPath $o2019OutputFolder -ShowUI $false -Version "O2019" -InstallPath $InstallerFolder
    Write-Host "Paths: $paths"

    Write-Host "Starting download Office 2019 with config xml file"
    $result = Start-Command -Path "$InstallerFolder\setup.exe" -Arguments "/download `"$($paths.ConfigPath)`""

    if ($result.ExitCode -eq 0) {
        Write-Host "Starting setup Office 2019"
        $result = Start-Command -Path "$InstallerFolder\setup.exe" -Arguments "/configure `"$($paths.ConfigPath)`""
    }

    [pscustomobject]@{
        LogPath  = $paths.LogPath
        ExitCode = $result.ExitCode
    }
}

function Get-LogList
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]
        $LogPath
    )

    $logList = New-Object System.Collections.ArrayList
    $lastLogItem = $null

    Get-Content $LogPath | %{
        if ($_ -match "^(\d{4}-\d{2}-\d{2}\s\d{2}:\d{2}:\d{2}\,\d{3})\s\[([A-Za-z0-9]+)\]\s(\w*)\s{1,5}([\u0000-\uFFFF]*)$") {
            $lastLogItem = [pscustomobject]@{
                Date  = [datetime]::parseexact($Matches[1], 'yyyy-MM-dd HH:mm:ss,fff', $null)
                ThreadName = $Matches[2]
                Type = $Matches[3]
                Message = $Matches[4]
                RawText = $Matches[0]
                Attached = New-Object System.Collections.ArrayList
            }

            $logList.Add($lastLogItem) | Out-Null
        }elseif($lastLogItem -ne $null){
            $lastLogItem.Attached.Add($_) | Out-Null
        }else { #unknown log format
            $logList.Add([pscustomobject]@{
                Date  = $null
                ThreadName = $null
                Type = "UNKNOWN"
                Message = $_
                RawText = $_
                Attached = New-Object System.Collections.ArrayList
            }) | Out-Null
        }
    }

    $logList
}

function Get-LogErrors
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]
        $LogPath
    )
    $errorLogList = New-Object System.Collections.ArrayList
    $errorLog = Get-LogList -LogPath $LogPath
    $errorLog | %{
        if ($_.Type -ieq "FATAL" -or $_.Type -ieq "ERROR") {
            $errorLogList.Add($_) | Out-Null
        }
    }

    $errorLogList
}

<#
 .Synopsis
  Download Office installer from artifactory and install it

 .DESCRIPTION
  Download Office installer from artifactory and install it

 .PARAMETER Version
  Office version, the value should be ("2013", "2019", "2013SP1").

 .PARAMETER ConnectionString
  Azure storage connection string.

 .EXAMPLE
   Install Office 2019.
   Install-Office -Version 2019 -ConnectionString "#########"

#>
function Install-Office {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]
        [ValidateSet("2013", "2019", "2013SP1")]
        $Version,
        [Parameter(Mandatory)]
        [string]
        $ConnectionString
    )

    $ctx = New-AzStorageContext -ConnectionString $ConnectionString

    $storagePath = "Office/$Version/"
    $installerDir = "$InstallerFolder\" + $($storagePath.Substring(0, $storagePath.Length - 1).Replace('/', '\'))

    Get-InstallPackage -StoragePath $storagePath -StorageContext $ctx -DestinationFolder "$InstallerFolder\" | Out-Null

    Write-Host "Office installer is in $installerDir"

    switch ($Version) {
        "2013" {  
            Install-Office2013 -InstallerFolder $installerDir
            break
        }
        "2013SP1" {  
            $kbStoragePath = "Office/2013SP1KBs/"
            $kbFolder = "$InstallerFolder\" + $($kbStoragePath.Substring(0, $kbStoragePath.Length - 1).Replace('/', '\'))
            Get-InstallPackage -StoragePath $kbStoragePath -StorageContext $ctx -DestinationFolder "$InstallerFolder\" | Out-Null
            Install-Office2013SP1 -InstallerFolder $installerDir -KBFolder $kbFolder
            break
        }
        "2019" { 
            Install-Office2019 -InstallerFolder $installerDir
            break
        }
        Default 
        { 
            throw "Office version is not correct"
        }
    }
}

<#
 .Synopsis
  Download ArchiveManager installer from artifactory and install it

 .DESCRIPTION
  Download ArchiveManager installer from artifactory and install it

 .PARAMETER Version
  ArchiveManager version e.g. 5.7.0.316

 .PARAMETER Username
  Artifactory credential username.

 .PARAMETER Password
  Artifactory credential password.

  .PARAMETER Branch
  Branch build to find the installer e.g. Trunk

 .PARAMETER Features
  Which ArchiveManager feature should be installed, this parameter can be null, e.g. @('ADC', 'ESM', 'Dataloader', 'FTI_FTS', 'Retention', 'Alert', 'Website')

 .EXAMPLE
   Install Archivemanager 5.7.0.316.
   Install-ArchiveManager -Version 5.7.0.316 -Username "user@prod.quest.corp" -Password "####" -Branch Trunk

#>

function Install-ArchiveManager {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]
        $Version,
        [Parameter(Mandatory)]
        [string]
        $Branch,
        [Parameter(Mandatory)]
        [string]
        $Username,
        [Parameter(Mandatory)]
        [string]
        $Password,
        [Parameter()]
        [string[]]
        $Features
    )

    $msi = "ArchiveManagerInstaller.msi"
    $msiFolder = "$InstallerFolder\$AMPrefix\$Version"
    $msiPath = "$msiFolder\$msi"
    $amOutputFolder = "$OutputFolder\$AMPrefix"
    $aBranch = $Branch.Trim(@('/'))

    $featureParam = ""
    if ($Features -ne $null -and $Features.Length -gt 0) {
        $featureParam = "FEATURELIST=" + [string]::Join(";", $($Features | %{$_ + "Feature"}))
    }

    New-Item -ItemType Directory -Force -Path $msiFolder | Out-Null

    $url = "https://artifactory.labs.quest.com/qam-build/builds/$aBranch/release/AM-$Version/$msi";

    $webClient = New-Object System.Net.WebClient
    $webClient.Credentials = New-Object System.Net.NetworkCredential($Username, $Password)
    Write-Host "Download file from `"$url`" to `"$msiPath`""
    $webClient.DownloadFile($url, $msiPath) | Write-Host

    New-Item -ItemType Directory -Force -Path $amOutputFolder | Out-Null

    $logFolder = Get-Date -Format "yyyyMMddHHmmss"
    $logPath = "$amOutputFolder\$logFolder"
    New-Item -ItemType Directory -Force -Path $logPath | Out-Null

    $result = Start-Command -Path "msiexec.exe" -Arguments "/i `"$msiPath`"  /qn /L*V `"$logPath\AMInstaller.log`" LICENSEACCEPTED=1 $featureParam"
    
    [pscustomobject]@{
        LogPath  = $logPath
        ExitCode = $result.ExitCode
    }
}

<#
.SYNOPSIS
Configure ArchiveManager Configuration Console Sliently, and add License

.DESCRIPTION
Configure ArchiveManager Configuration Console Sliently, and add License

.PARAMETER ConfigurationFile
The path of configuration xml file

.PARAMETER LicenseFile
The path of license file

.EXAMPLE
Set-ArchiveManager -ConfigurationFile "C:\Configuration.xml" -LicenseFile "C:\license.asc"

.NOTES
Return the psobject, there are 2 properties: 
1, Errors, the error log item list; 
    The log item define:
        [DateTime]Date,
        [string]ThreadName,
        [string]Type,
        [string]Message,
        [string]RawText,
        [ArrayList]Attached
2, ExitCode, ConfigConsoleSlient.exe exit code, 0 is succeed
#>

function Config-ArchiveManager
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]
        $ConfigurationFile,
        [Parameter(Mandatory)]
        [string]
        $LicenseFile
    )

    $path = Get-AMInstallationPath

    $configSlient = [System.IO.Path]::Combine($path, "ConfigurationConsoleSilent.exe")
    if (-not $(Test-Path $configSlient -PathType Leaf)) {
        throw "Cannot find the file $configSlient"
    }

    $logFolder = "C:\Quest\ArchiveManager\Logs"
    $regKey = $null

    if (Test-Path -Path "HKLM:\SOFTWARE\AfterMail") {
        $regKey = Get-ItemProperty -Path "HKLM:\SOFTWARE\AfterMail"
    }
    elseif (Test-Path -Path "HKLM:\SOFTWARE\Wow6432Node\AfterMail") {
        $regKey = Get-ItemProperty -Path "HKLM:\SOFTWARE\Wow6432Node\AfterMail"
    }

    if ($regKey -ne $null -and "Log Path" -in $regKey.PSobject.Properties.Name) {
        $logFolder = $($regKey | Get-ItemPropertyValue -Name "Log Path")
    }

    Write-Host "ArchiveManager log path is `"$logFolder`""

    $commandResult = Start-Command -Path $configSlient -Arguments "-configFile `"$ConfigurationFile`" -licenseFile `"$LicenseFile`"" -UseShell

    $errors = Get-LogErrors -LogPath $([System.IO.Path]::Combine($logFolder, "ConfigConsoleSilent.wlog"))
    
    [pscustomobject]@{
        Errors  = $errors
        ExitCode = $commandResult.ExitCode
    }
}

