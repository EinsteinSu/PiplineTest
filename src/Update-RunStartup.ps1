param(
    [Parameter(Mandatory)]
    [string]
    $BaseFolder
)
$resourceStorageAccount = Get-AzStorageAccount -ResourceGroupName "automationlabs" `
                                -Name "automationadmin"
$testStorageContext = $resourceStorageAccount.Context;
$fileName = "Run-Startup.ps1";
$file = [System.IO.Path]::Combine($BaseFolder, "src", $fileName)
Set-AzStorageBlobContent -File $file -Container "startup" -Blob $fileName -Context $testStorageContext -Force;