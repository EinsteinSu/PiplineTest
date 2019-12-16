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
    [ValidateSet("ExchangeOnline","SingleExchange","Groupwise")]
    [string]
    $Environment,
    [string]
    $Dns,
    [string]
    [Parameter(Mandatory)]
	$TestResourceGroupName,
    [Parameter(Mandatory)]
    [string]
    $ResourceStorageAccountName,
    [Parameter(Mandatory)]
    [string]
    $ResourcesContainerName,
    [Parameter(Mandatory)]
    [string]
    $Tenant,
    [Parameter(Mandatory)]
    [string]
    $TestResultContainerName
)

$baseFolder = "C:\AutomationTest\"
New-Item -Path $baseFolder -ItemType "directory" -Force

Write-Host "Logging in to Azure";
$cred =  New-Object System.Management.Automation.PSCredential ($AzAccount,(ConvertTo-SecureString $AzPassword -AsPlainText -Force)) 
Connect-AzAccount -ServicePrincipal -Tenant $Tenant -Credential $cred

Write-Host "Downloading resource files from $ResourcesContainerName to $baseFolder"
$resourceStorageAccount = Get-AzStorageAccount -ResourceGroupName $TestResourceGroupName `
                                -Name $ResourceStorageAccountName

$ctx = $resourceStorageAccount.Context;
Get-AzStorageBlob -Container $ResourcesContainerName -Context $ctx | Get-AzStorageBlobContent -Destination $baseFolder -Force

Set-Location -Path $baseFolder
.\Startup.ps1 -AzAccount $AzAccount `
            -AzPassword $AzPassword `
            -StorageAccountName $StorageAccountName `
            -StorageKey $StorageKey `
            -StorageConnection $StorageConnection `
            -OutlookVersion $OutlookVersion `
            -QamVersion $QamVersion `
            -Branch $Branch `
            -ArtUserName $ArtUserName `
            -ArtPassword $ArtPassword `
            -TestComponent $TestComponent `
            -TestTags $TestTags `
            -InstallFeatures $InstallFeatures `
            -Dns $Dns `
            -TestResourceGroupName $TestResourceGroupName `
            -Environment $Environment `
            -Tenant $Tenant `
            -TestResultContainerName $TestResultContainerName `
            -BaseFolder $baseFolder