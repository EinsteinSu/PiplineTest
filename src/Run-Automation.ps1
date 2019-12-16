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
    $OutlookVersions,
    [string]
    $QamVersion,
    [string]
    $Branch,
    [string]
    $ArtUserName,
    [string]
    $ArtPassword,
    [ValidateSet("ADC", "ESM", "MAPIDL", "FTI", "FTS", "GDC", "GSM", "Retention", "Alert", "")]
    [string]
    $TestComponent,
    [string]
    $TestTags,
    [string]
    $InstallFeatures,
    [string]
    $TestResourceGroupName,
    [ValidateSet("ExchangeOnline","SingleExchange","Groupwise")]
    [string]
    $Environment,
    [Parameter(Mandatory)]
    [string]
    $ResourcePath,
    [Parameter(Mandatory)]
    [string]
    $Location,
    [Parameter(Mandatory)]
    [string]
    $DCTag,
    [Parameter(Mandatory)]
    [string]
    $QAMTag,
    [Parameter(Mandatory)]
    [string]
    $VMSize,
    [Parameter(Mandatory)]
    [string]
    $Tenant,
    [Parameter(Mandatory)]
    [string]
    $DnsServerAddress,
    [Parameter(Mandatory)]
    [string]
    $QamServerAddress,
    [Parameter(Mandatory)]
    [string]
    $ResourcesContainerName,
    [Parameter(Mandatory)]
    [string]
    $TestResultContainerName
)
$ErrorActionPreference = "Stop"

$groupName = "AutomationLabs";
$nsgName = "dc-nsg";
$storageConnection = "DefaultEndpointsProtocol=https;AccountName=$storageAccountName;AccountKey=$storageKey;EndpointSuffix=core.windows.net";


$resourceStorageAccountName = $TestResourceGroupName.ToLower() + "storages";
Write-Host "Creating azure storage $resourceStorageAccountName";
$resourceStorageAccount = New-AzStorageAccount -ResourceGroupName $TestResourceGroupName `
                                -Name $resourceStorageAccountName `
                                -SkuName Standard_LRS `
                                -Location $Location;
$ctx = $resourceStorageAccount.Context;

Write-Host "Creating Container $ResourcesContainerName";
New-AzStorageContainer -Name $ResourcesContainerName -Context $ctx -Permission blob;

New-AzStorageContainer -Name $TestResultContainerName -Context $ctx -Permission blob;
$ResourcePath = [System.IO.Path]::Combine($ResourcePath, "src")
Write-Host "Uploading file from $ResourcePath to $resourceStorageAccount - $ResourcesContainerName";
Get-ChildItem -File $ResourcePath -Recurse | Set-AzStorageBlobContent -Context $ctx -Container $ResourcesContainerName;

function Get-ExecutionCommand($Name, $Value){
    if([String]::IsNullOrEmpty($Value)){
        return "";
    }
    Write-Host "-$Name $Value"
    return "-$Name $Value ";
}

function Get-ExtensionCommand($OutlookVersion){
    $command = "powershell -ExecutionPolicy Unrestricted -File run-startup.ps1 ";
    $command += Get-ExecutionCommand -Name "StorageAccountName" -Value $StorageAccountName;
    $command += Get-ExecutionCommand -Name "StorageKey" -Value $StorageKey;
    $command += Get-ExecutionCommand -Name "StorageConnection" -Value $StorageConnection;
    $command += Get-ExecutionCommand -Name "OutlookVersion" -Value $OutlookVersion;
    $command += Get-ExecutionCommand -Name "QamVersion" -Value $QamVersion;
    $command += Get-ExecutionCommand -Name "Branch" -Value $branch;
    $command += Get-ExecutionCommand -Name "ArtUserName" -Value $ArtUserName;
    $command += Get-ExecutionCommand -Name "ArtPassword" -Value $ArtPassword;
    $command += Get-ExecutionCommand -Name "TestComponent" -Value $TestComponent;
    $command += Get-ExecutionCommand -Name "TestTags" -Value $TestTags;
    $command += Get-ExecutionCommand -Name "InstallFeatures" -Value $InstallFeatures;
    $command += Get-ExecutionCommand -Name "AzAccount" -Value $AzAccount;
    $command += Get-ExecutionCommand -Name "AzPassword" -Value $AzPassword;
    $command += Get-ExecutionCommand -Name "Environment" -Value $Environment;
    $command += Get-ExecutionCommand -Name "Tenant" -Value $Tenant;
    $command += Get-ExecutionCommand -Name "Dns" -Value $DnsServerAddress;
    $command += Get-ExecutionCommand -Name "TestResourceGroupName" -Value $TestResourceGroupName;
    $command += Get-ExecutionCommand -Name "ResourceStorageAccountName" -Value $ResourceStorageAccountName;
    $command += Get-ExecutionCommand -Name "ResourcesContainerName" -Value $ResourcesContainerName;
    $command += Get-ExecutionCommand -Name "TestResultContainerName" -Value $TestResultContainerName;
    return $command;
}

function New-VM($VmName, $SnapshotName, $IpAddress, $Vnet, $OutlookVersion){
    
    Write-Output "Creating $VmName";
    $diskName = "$SnapshotName" +"_" + $OutlookVersion + "_copy";
    Write-Output "Getting the snapshot $SnapshotName";
    $snapshot = Get-AzSnapshot $groupName -SnapshotName $SnapshotName;;

    Write-Output "Creating disk $diskName"
    $disConfig = New-AzDiskConfig -Location $snapshot.Location -SourceResourceId $snapshot.Id -CreateOption Copy;
    $disk = New-AzDisk -Disk $disConfig -ResourceGroupName $testResourceGroupName -DiskName $diskName;

    Write-Output "Setting the vm size to $VMSize";
    $vm = New-AzVMConfig -VMName $VmName -VMSize $VMSize;
    $vm = Set-AzVMOSDisk -VM $vm -ManagedDiskId $disk.Id -CreateOption Attach -Windows;

    Write-Output "Creating network for $VmName";
    $publicIpName = ($VmName.ToLower()+'_ip');
    $publicIp = New-AzPublicIpAddress -Name $publicIpName -ResourceGroupName $testResourceGroupName -Location $snapshot.Location -AllocationMethod Dynamic;
    $privateIpName = ($VmName.ToLower()+'_pip');
    New-AzNetworkInterfaceIpConfig -Name $privateIpName -Subnet $Vnet.Subnets[0] -PrivateIpAddress $IpAddress -Primary;
    $nicName = ($VmName.ToLower()+'_nic');
    $nic = New-AzNetworkInterface -Name $nicName -ResourceGroupName $testResourceGroupName -Location $snapshot.Location -SubnetId $vnet.Subnets[0].Id -IpConfigurationName $privateIpName -PublicIpAddressId $publicIp.Id;
    $vm = Add-AzVMNetworkInterface -VM $vm -Id $nic.Id;

    Write-Output "Creating VM $VmName";
    New-AzVM -VM $vm -ResourceGroupName $testResourceGroupName -Location $snapshot.Location;
}



$tags = @{ Config = $DCTag };
$dcSnapshotName = (Get-AzResource -Tag $tags).Name;
if($null -eq $dcSnapshotName){
    Write-Error "Can not found snapshot for DC with $DCTag";
    Exit-PSSession;
}
$tags = @{ Config = $QAMTag };
$qamSnapshotName = (Get-AzResource -Tag $tags).Name;
if($null -eq $qamSnapshotName){
    Write-Error "Can not found snapshot for QAM with $QAMTag";
    Exit-PSSession;
}

foreach($outlookVersion in $OutlookVersions.Split(',')){
    $outlookVersion = $outlookVersion.Trim();
    Write-Host "Testing for Outlook $outlookVersion in $exchangeVersion $os $dbVersion";
    $securityGroup = Get-AzNetworkSecurityGroup -Name $nsgName -ResourceGroupName $groupName;
    $subnetConfigName = "subnet_$outlookVersion";
    $subnetConfig = New-AzVirtualNetworkSubnetConfig -Name $subnetConfigName -AddressPrefix 172.31.11.0/24 -NetworkSecurityGroup $securityGroup;
     
    $vitrualNetworkName = "vmNetwork_$outlookVersion";
    Write-Output "Creating the virtual network $vitrualNetworkName";
    $vnet = New-AzVirtualNetwork -Name $vitrualNetworkName -ResourceGroupName $testResourceGroupName -Location $Location -AddressPrefix 172.31.0.0/16 -Subnet $subnetConfig;

    $vmDcName = "dc$outlookVersion";
    $vmQAMName = "qam$outlookVersion";
    New-VM -VmName $vmDcName -SnapshotName $dcSnapshotName -IpAddress $dnsServerAddress -Vnet $vnet -OutlookVersion $outlookVersion;
    New-VM -VmName $vmQAMName -SnapshotName $qamSnapshotName -IpAddress $QamServerAddress -Vnet $vnet -OutlookVersion $outlookVersion;

    $command = Get-ExtensionCommand -outlookVersion $outlookVersion
    Write-Host "Start getting the startup script to install QAM"
    $fileUri = @("https://automationadmin.blob.core.windows.net/startup/Run-Startup.ps1")
    $settings = @{"fileUris" = $fileUri};
    $protectedSettings = @{"storageAccountName" = $storageAccountName; "storageAccountKey" = $storageKey; "commandToExecute" = $command};
    $extensionName ="InstallQAMFor$outlookVersion";
    Write-Host "Executing the extension Command: $command"
    Set-AzVMExtension -ResourceGroupName $testResourceGroupName `
        -Location $Location `
        -VMName $vmQAMName `
        -Name $extensionName `
        -Publisher "Microsoft.Compute" `
        -ExtensionType "CustomScriptExtension" `
        -TypeHandlerVersion "1.9" `
        -Settings $settings    `
        -ProtectedSettings $protectedSettings

    $output = Get-AzVMDiagnosticsExtension -ResourceGroupName $testResourceGroupName -VMName $vmQAMName -Name $extensionName -Status #-Debug
    $text = $output.SubStatuses[0].Message
    [regex]::Replace($text, "\\n", "`n")
}