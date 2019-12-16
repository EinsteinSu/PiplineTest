param(
	[Parameter(Mandatory)]
    [string]
    $AzAccount,
	[Parameter(Mandatory)]
    [string]
    $AzPassword,
    [Parameter(Mandatory)]
    [string]
    $StorageAccountName,
    [Parameter(Mandatory)]
    [string]
    $StorageKey,
    [string]
    $StorageConnection,
    [Parameter(Mandatory)]
    [string]
    $OutlookVersion,
    [Parameter(Mandatory)]
    [string]
    $QamVersion,
    [Parameter(Mandatory)]
    [string]
    $Branch,
    [Parameter(Mandatory)]
    [string]
    $ArtUserName,
    [Parameter(Mandatory)]
    [string]
    $ArtPassword,
    [string]
    $TestComponent,
    [string]
    $TestTags,
    [string]
    $InstallFeatures,
    [Parameter(Mandatory)]
    [string]
    $Dns,
    [Parameter(Mandatory)]
    [string]
    $TestResourceGroupName,
    [Parameter(Mandatory)]
    [string]
    $ResourceStorageAccountName,
    [ValidateSet("ExchangeOnline","SingleExchange","Groupwise")]
    [string]
    $Environment,
    [Parameter(Mandatory)]
    [string]
    $Tenant,
    [Parameter(Mandatory)]
    [string]
    $TestResultContainerName,
    [Parameter(Mandatory)]
    [string]
    $BaseFolder
)

function Write-ExecutionLog($Result, $ExecutionName){
    if($Result.ExitCode -eq 0){
        Write-Host "Successfully $ExecutionName";
    }else{
        Write-Host "Failed to $ExecutionName";
        $items = Get-ChildItem -Path $Result.LogPath;
        foreach($item in $items){
            $text = Get-Content -Path $item.FullName;
            Write-Host $text;
        }
    }
}

Write-Host "Setting Dns to $Dns";
$net = Get-NetIPConfiguration | Select-Object InterfaceIndex;
$index = $net[0];
Set-DnsClientServerAddress -InterfaceIndex $index.InterfaceIndex -ServerAddresses $Dns

$resourceGroup = "AutomationLabs";
$storageAccount = Get-AzStorageAccount -Name $StorageAccountName -ResourceGroupName $resourceGroup;
$ctx = $storageAccount.Context;
$container = "startup";
$installationFolder = "C:\Installer\";
New-Item -Path $installationFolder -ItemType "directory" -Force
Write-Host "Downloading files from $container";
Get-AzStorageBlobContent -Blob "license.asc" -Container $container -Destination $installationFolder -Context $ctx -Force


Write-Host "Import modules AZ and Installer"
Import-Module AZ
Import-Module $BaseFolder\Modules\Installer\Installer.psd1 -Force;

$result = Install-Office -Version $OutlookVersion -ConnectionString $StorageConnection;
Write-ExecutionLog -Result $result -ExecutionName "Installed Office";

$result = Install-ArchiveManager -Version $QamVersion -Branch $Branch -Username $ArtUserName -Password $ArtPassword;
Write-ExecutionLog -Result $result -ExecutionName "Installed Quest Archive Manager";


$configFile = "$BaseFolder\Configurations\$Environment\Configuration.xml";
$licenseFile = [System.IO.Path]::Combine($installationFolder, "license.asc")
Write-Host "Configure QAM with file $configFile $licenseFile";
$result = Config-ArchiveManager -ConfigurationFile $configFile -LicenseFile $licenseFile;
Write-ExecutionLog -Result $result -ExecutionName "Configured Quest Archive Manager";

Set-Location -Path $BaseFolder
$result = .\Run-Test.ps1;
if($result){
    Write-Error "The Automation tests executing failed.";
}else{
    Write-Host "The Automation tests Successfully executed, get the test result file from $result.";
}


$resourceStorageAccount = Get-AzStorageAccount -ResourceGroupName $TestResourceGroupName `
                                -Name $ResourceStorageAccountName
$testStorageContext = $resourceStorageAccount.Context;
Set-AzStorageBlobContent -File $result -Container $TestResultContainerName -Blob $OutlookVersion + "_testresult.xml" -Context $testStorageContext -Force;
