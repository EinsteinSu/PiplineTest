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
    $OutlookVersions,
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
    [Parameter()]
    [string]
    $TestResourceGroupName
)

$groupName = "AutomationLabs"
$vitrualNetworkName = "vmNetwork_" + $batch
$vmSize = "Standard_DS3"
$location = "WestUS"
$nsgName = "dc-nsg"
$exchangeVersion = "Ex2019_CU3";
$os = "Win2019"
$dbVersion = "SQL2014"
$storageConnection = "DefaultEndpointsProtocol=https;AccountName=$storageAccountName;AccountKey=$storageKey;EndpointSuffix=core.windows.net";
function Get-ExecutionCommand($Name, $Value){
    if([String]::IsNullOrEmpty($Value)){
        return "";
    }
    Write-Host "-$Name $Value"
    return "-$Name $Value ";
}

function Get-ExtensionCommand($outlookVersion){
    $command = "powershell -ExecutionPolicy Unrestricted -File startup.ps1 ";
    $command += Get-ExecutionCommand -Name "StorageAccountName" -Value $storageAccountName;
    $command += Get-ExecutionCommand -Name "StorageKey" -Value $storageKey;
    $command += Get-ExecutionCommand -Name "StorageConnection" -Value $storageConnection;
    $command += Get-ExecutionCommand -Name "OutlookVersion" -Value $outlookVersion;
    $command += Get-ExecutionCommand -Name "QamInstallerVersion" -Value $qamInstallerVersion;
    $command += Get-ExecutionCommand -Name "Branch" -Value $branch;
    $command += Get-ExecutionCommand -Name "ArtUserName" -Value $artUserName;
    $command += Get-ExecutionCommand -Name "ArtPassword" -Value $artPassword;
    $command += Get-ExecutionCommand -Name "TestComponent" -Value $testComponent;
    $command += Get-ExecutionCommand -Name "TestTags" -Value $testTags;
    $command += Get-ExecutionCommand -Name "InstallFeatures" -Value $installFeatures;
    $command += Get-ExecutionCommand -Name "AzAccount" -Value $azAccount;
    $command += Get-ExecutionCommand -Name "AzPassword" -Value $azPassword;
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

    Write-Output "Setting the vm size to $vmSize";
    $vm = New-AzVMConfig -VMName $VmName -VMSize $vmSize;
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



$tags = [ordered]@{Type = "DC";  ExchangeVersion = $exchangeVersion; OS = $os}
$dcSnapshotName = (Get-AzResource -Tag $tags)[0].Name;
if($null -eq $dcSnapshotName){
    Write-Error "Can not found snapshot for DC";
    Exit-PSSession;
}
$tags = [ordered]@{Type = "QAM"; DbVersion = $dbVersion ; OS = $os;}
$qamSnapshotName = (Get-AzResource -Tag $tags)[0].Name;
if($null -eq $qamSnapshotName){
    Write-Error "Can not found snapshot for QAM";
    Exit-PSSession;
}

foreach($outlookVersion in $OutlookVersions.Split(',')){
    $outlookVersion = $outlookVersion.Trim();
    Write-Host "Testing for Outlook $outlookVersion in $exchangeVersion $os $dbVersion";
    $securityGroup = Get-AzNetworkSecurityGroup -Name $nsgName -ResourceGroupName $groupName;
    $subnetConfigName = "subnet_" + $batch;
    $subnetConfig = New-AzVirtualNetworkSubnetConfig -Name $subnetConfigName -AddressPrefix 172.31.11.0/24 -NetworkSecurityGroup $securityGroup;
     
    
    Write-Output "Creating the virtual network"
    $vnet = New-AzVirtualNetwork -Name $vitrualNetworkName -ResourceGroupName $testResourceGroupName -Location $location -AddressPrefix 172.31.0.0/16 -Subnet $subnetConfig;
    
    $vmDcName = "dc$outlookVersion";
    $vmQAMName = "qam$outlookVersion";
    New-VM -VmName $vmDcName -SnapshotName $dcSnapshotName -IpAddress "172.31.11.5" -Vnet $vnet -OutlookVersion $outlookVersion;
    New-VM -VmName $vmQAMName -SnapshotName $qamSnapshotName -IpAddress "172.31.11.4" -Vnet $vnet -OutlookVersion $outlookVersion;

    $command = Get-ExtensionCommand -outlookVersion $outlookVersion
    Write-Host "Start getting the startup script to install QAM"
    $fileUri = @("https://automationadmin.blob.core.windows.net/startup/startup.ps1")
    $settings = @{"fileUris" = $fileUri};
    $protectedSettings = @{"storageAccountName" = $storageAccountName; "storageAccountKey" = $storageKey; "commandToExecute" = $command};
    $extensionName ="InstallQAMFor$outlookVersion";
    Write-Host "Executing the extension Command: $command"
    Set-AzVMExtension -ResourceGroupName $testResourceGroupName `
        -Location $location `
        -VMName $vmQAMName `
        -Name $extensionName `
        -Publisher "Microsoft.Compute" `
        -ExtensionType "CustomScriptExtension" `
        -TypeHandlerVersion "1.9" `
        -Settings $settings    `
        -ProtectedSettings $protectedSettings
}


 






