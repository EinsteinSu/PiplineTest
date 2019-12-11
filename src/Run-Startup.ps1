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
    [ValidateSet("ExchangeOnline","SingleExchange","Groupwise")]
    [string]
    $Environment,
    [Parameter()]
    [string]
    $Dns,
    [Parameter()]
    [string]
    $TestResourceGorupName,
    [Parameter(Mandatory)]
    [string]
    $ResourceStorageAccountName,
    [Parameter(Mandatory)]
    [string]
    $ResourceStorageContainerName
)


Write-Host "Logging in to Azure"
$tenant = "91c369b5-1c9e-439c-989c-1867ec606603";
$cred =  New-Object System.Management.Automation.PSCredential ($AzAccount,(ConvertTo-SecureString $AzPassword -AsPlainText -Force)) 
Connect-AzAccount -ServicePrincipal -Tenant $tenant -Credential $cred

New-Item -Path "C:\" -ItemType "directory" -Name "AutomationTest" -Force
$baseFolder = "C:\AutomationTest"

Write-Host "Downloading resource files from $ResourceStorageContainerName to $baseFolder"
$resourceStorageAccount = Get-AzStorageAccount -ResourceGroupName $TestResourceGroupName `
                                -Name $ResourceStorageAccountName `

$ctx = $resourceStorageAccount.Context;
Get-AzStorageBlob -Container $ResourceStorageContainerName -Context $ctx | Get-AzStorageBlobContent -Destination $baseFolder -Force

C:\Startup.ps1 -AzAccount $AzAccount `
            -AzPassword $AzPassword `
            -StorageAccountName $StorageAccountName `
            -StorageKey $StorageKey `
            -StorageConnection $StorageConnection `
            -OutlookVersion $OutlookVersion `
            -QamInstallerVersion $QamInstallerVersion `
            -Branch $Branch `
            -ArtUserName $ArtUserName `
            -ArtPassword $ArtPassword `
            -TestComponent $TestComponent `
            -TestTags $TestTags `
            -InstallFeatures $InstallFeatures `
            -Dns $Dns `
            -TestResourceGorupName $TestResourceGorupName `
            -Environment $Environment
