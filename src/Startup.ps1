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
    $TestResourceGroupName,
    [Parameter(Mandatory)]
    [string]
    $ResourceStorageAccountName,
    [ValidateSet("ExchangeOnline","SingleExchange","Groupwise")]
    [string]
    $Environment
)
Write-Host "Setting Dns to $Dns";
$net = Get-NetIPConfiguration | Select-Object InterfaceIndex;
$index = $net[0];
Set-DnsClientServerAddress -InterfaceIndex $index.InterfaceIndex -ServerAddresses $Dns

Write-Host "Logging in to Azure";
$tenant = "91c369b5-1c9e-439c-989c-1867ec606603";
$cred =  New-Object System.Management.Automation.PSCredential ($AzAccount,(ConvertTo-SecureString $AzPassword -AsPlainText -Force)) 
Connect-AzAccount -ServicePrincipal -Tenant $tenant -Credential $cred

$resourceGroup = "AutomationLabs";
$storageAccount = Get-AzStorageAccount -Name $StorageAccountName -ResourceGroupName $resourceGroup;
$ctx = $storageAccount.Context;
$container = "startup";
$installationFolder = "C:\Installer\";
New-Item -Path "C:\" -ItemType "directory" -Name "Installer" -Force
Write-Host "Downloading files from $container";
Get-AzStorageBlobContent -Blob "license.asc" -Container $container -Destination $installationFolder -Context $ctx -Force


$baseFolder = "C:\AutomationTest"
Write-Host "Import modules AZ and Installer"
Import-Module AZ
Import-Module C:\AutomationTest\Modules\Installer\Installer.psd1 -Force;


function Write-ExecutionLog($Result, $ExecutionName){
    if($Result.ExitCode -eq 0){
        Write-Host "Successfully $ExecutionName office.";
    }else{
        $items = Get-ChildItem -Path $Result.LogPath;
        foreach($item in $items){
            $text = Get-Content -Path $item.FullName;
            Write-Host $text;
        }
    }
}

$result = Install-Office -Version $OutlookVersion -ConnectionString $StorageConnection;
Write-ExecutionLog -Result $result -ExecutionName "Installed Office";

$result = Install-ArchiveManager -Version $QamInstallerVersion -Branch $Branch -Username $ArtUserName -Password $ArtPassword;
Write-ExecutionLog -Result $result -ExecutionName "Installed Quest Archive Manager";


$configFile = "$baseFolder\Configure\$Environment\Configuration.xml";
$licenseFile = $installationFolder + "license.asc";
Write-Host "Configure QAM with file $configFile $licenseFile";
$result = Config-ArchiveManager -ConfigurationFile $configFile -LicenseFile $licenseFile;
Write-ExecutionLog -Result $result -ExecutionName "Configured Quest Archive Manager";

$result = C:\AutomationTest\Run-Test.ps1;
if($result){
    Write-Error "The Automation tests executing failed.";
}else{
    Write-Host "The Automation tests Successfully executed, get the test result file from $result.";
}


$resourceStorageAccount = Get-AzStorageAccount -ResourceGroupName $TestResourceGroupName `
                                -Name $ResourceStorageAccountName
$testStorageContext = $resourceStorageAccount.Context;
Set-AzStorageBlobContent -File $result -Container $containerName -Blob "testresult.xml" -Context $testStorageContext -Force;
