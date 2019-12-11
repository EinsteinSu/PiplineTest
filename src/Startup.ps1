param(
	[Parameter()]
    [string]
    $AzAccount,
	[Parameter()]
    [string]
    $AzPassword,
    [Parameter()]
    [string]
    $StorageAccountName,
    [Parameter()]
    [string]
    $StorageKey,
    [Parameter()]
    [string]
    $StorageConnection,
    [Parameter()]
    [string]
    $OutlookVersion,
    [Parameter()]
    [string]
    $QamInstallerVersion,
    [Parameter()]
    [string]
    $Branch,
    [Parameter()]
    [string]
    $ArtUserName,
    [Parameter()]
    [string]
    $ArtPassword,
    [Parameter()]
    [string]
    $TestComponent,
    [Parameter()]
    [string]
    $TestTags,
    [Parameter()]
    [string]
    $InstallFeatures,
    [Parameter(Mandatory)]
    [string]
    $Dns,
    [Parameter(Mandatory)]
    [string]
    $TestResourceGorupName,
    [ValidateSet("ExchangeOnline","SingleExchange","Groupwise")]
    [string]
    $Environment
)
Write-Host "Setting Dns to $Dns"
$net = Get-NetIPConfiguration | Select-Object InterfaceIndex;
$index = $net[0];
Set-DnsClientServerAddress -InterfaceIndex $index.InterfaceIndex -ServerAddresses $Dns

Write-Host "Logging in to Azure"
$tenant = "91c369b5-1c9e-439c-989c-1867ec606603";
$cred =  New-Object System.Management.Automation.PSCredential ($AzAccount,(ConvertTo-SecureString $AzPassword -AsPlainText -Force)) 
Connect-AzAccount -ServicePrincipal -Tenant $tenant -Credential $cred

$resourceGroup = "AutomationLabs";
$storageAccount = Get-AzStorageAccount -Name $StorageAccountName -ResourceGroupName $resourceGroup;
$ctx = $storageAccount.Context;
$container = "startup"
$installationFolder = "C:\Installer\";
New-Item -Path "C:\" -ItemType "directory" -Name "Installer" -Force
Get-AzStorageBlobContent -Blob "license.asc" -Container $container -Destination $installationFolder -Context $ctx -Force
#copy the config file to C:\Installer


#import modules
Import-Module AZ
Import-Module C:\Installer\Installer.psd1

function Write-ExecutionLog($Result, $ExecutionName){
    if($Result.ExitCode -eq 0){
        Write-Host "Successfully $ExecutionName office.";
    }else{
        $items = Get-ChildItem -Path $Result.LogPath
        foreach($item in $items){
            Get-Content -Path $item.FullName
        }
    }
}

$result = Install-Office -Version $OutlookVersion -ConnectionString $StorageConnection
Write-ExecutionLog -Result $result -ExecutionName "Office"

$result = Install-ArchiveManager -Version $QamInstallerVersion -Branch $Branch -Username $ArtUserName -Password $ArtPassword
Write-ExecutionLog -Result $result -ExecutionName "Quest Archive Manager"


$configFile = $installationFolder + "Configuration.xml"
$licenseFile = $installationFolder + "license.asc"
$result = Config-ArchiveManager -ConfigurationFile $configFile -LicenseFile $licenseFile
Write-ExecutionLog -Result $result -ExecutionName "Configure Quest Archive Manager"


$atFolder = "C:\AT\"
New-Item -Path "C:\" -ItemType "directory" -Name "AT" -Force
Get-AzStorageBlob -Container "automationtest" -Context $ctx | Get-AzStorageBlobContent -Destination $atFolder -Force

$result = C:\AT\Run-Test.ps1
Write-Host "Get the test result file $result"
Set-AzStorageBlobContent -File $result -Container "startup" -Blob "testresult.xml" -Context $ctx -Force
