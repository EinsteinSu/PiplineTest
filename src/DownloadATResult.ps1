param(
    [Parameter(Mandatory)]
    [string]
    $TestResourceGroupName,
    [Parameter(Mandatory)]
    [string]
    $Destination,
    [Parameter(Mandatory)]
    [string]
    $TestResultContainerName
)

$resourceStorageAccountName = "$TestResourceGroupName".ToLower() + "storages";
Write-Host "Downloading test result files from $TestResultContainerName to $Destination"
$resourceStorageAccount = Get-AzStorageAccount -ResourceGroupName $TestResourceGroupName `
                                -Name $resourceStorageAccountName
$ctx = $resourceStorageAccount.Context;
Get-AzStorageBlob -Container $TestResultContainerName -Context $ctx | Get-AzStorageBlobContent -Destination $Destination -Force;