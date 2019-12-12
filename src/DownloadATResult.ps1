param(
    [Parameter(Mandatory)]
    [string]
    $TestResourceGroupName,
    [Parameter(Mandatory)]
    [string]
    $Destination
)
$resourceStorageAccountName = "$TestResourceGroupName".ToLower() + "storages";
$testResultContainer = "testresult";
Write-Host "Downloading test result files from $testResultContainer to $Destination"
$resourceStorageAccount = Get-AzStorageAccount -ResourceGroupName $TestResourceGroupName `
                                -Name $resourceStorageAccountName
$ctx = $resourceStorageAccount.Context;
Get-AzStorageBlob -Container $testResultContainer -Context $ctx | Get-AzStorageBlobContent -Destination $Destination -Force