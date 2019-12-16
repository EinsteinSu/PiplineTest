param(
    [Parameter(Mandatory)]
    [string]
    $BaseFolder
)
$resourceStorageAccount = Get-AzStorageAccount -ResourceGroupName "automationlabs" `
                                -Name "automationadmin"
$testStorageContext = $resourceStorageAccount.Context;
$file = [System.IO.Path]::Combine($BaseFolder, "src\Run-Startup.ps1")
Set-AzStorageBlobContent -File $file -Container "startup" -Blob "testresult.xml" -Context $testStorageContext -Force;