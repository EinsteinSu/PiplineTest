param(
	[Parameter()]
    [string]
    $FileName,
	[Parameter()]
    [string]
    $Destination,
    [string]
    $StorageAccountName
)
$resourceGroup = "AutomationLabs";
$storageAccount = Get-AzStorageAccount -Name $StorageAccountName -ResourceGroupName $resourceGroup;
$ctx = $storageAccount.Context;
$container = "startup"
$Destination = $Destination + "\"
New-Item -Path $Destination -ItemType "directory" -Name "ATTestResults" -Force
Write-Host "Copying $FileName to $Destination"
Get-AzStorageBlobContent -Blob $FileName -Container $container -Destination $Destination -Context $ctx -Force
