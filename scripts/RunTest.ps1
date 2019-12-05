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
    [Parameter()]
    [string]
    $TestResourceGroupName
)
$groupName = "AutomationLabs"
$vitrualNetworkName = "vmNetwork_" + $batch
$vmSize = "Standard_DS3"
$location = "WestUS"
$nsgName = "dc-nsg"

Write-Host "Outlook versin $OutlookVersion"

$exchangeVersion = "Ex2019_CU3";
$os = "Win2019"
$dbVersion = "SQL2014"

function Get-ExecutionCommand($Name, $Value){
    if([String]::IsNullOrEmpty($Value)){
        return "";
    }
    Write-Host "-$Name $Value"
    return "-$Name $Value ";
}
$storageConnection = "DefaultEndpointsProtocol=https;AccountName=$storageAccountName;AccountKey=$storageKey;EndpointSuffix=core.windows.net";

<#Write-Host "Log in Azure"
$tenant = "91c369b5-1c9e-439c-989c-1867ec606603";
$cred =  New-Object System.Management.Automation.PSCredential ($AzAccount,(ConvertTo-SecureString $AzPassword -AsPlainText -Force)) 
Connect-AzAccount -ServicePrincipal -Tenant $tenant -Credential $cred#>

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

Write-Output "Creating group $testResourceGroupName";
New-AzResourceGroup -Name $testResourceGroupName -Location $location;
 

$securityGroup = Get-AzNetworkSecurityGroup -Name $nsgName -ResourceGroupName $groupName;
$subnetConfigName = "subnet_" + $batch;
$subnetConfig = New-AzVirtualNetworkSubnetConfig -Name $subnetConfigName -AddressPrefix 172.31.11.0/24 -NetworkSecurityGroup $securityGroup;
 

Write-Output "Creating the virtual network"
$vnet = New-AzVirtualNetwork -Name $vitrualNetworkName -ResourceGroupName $testResourceGroupName -Location $location -AddressPrefix 172.31.0.0/16 -Subnet $subnetConfig;

 
function New-VM($vmName, $snapshotName, $ipAddress){
    Write-Output "Creating $vmName";
    $diskName = "$snapshotName" +"_" +$batch + "_copy";
    Write-Output "Getting the snapshot $snapshotName";
    $snapshot = Get-AzSnapshot $groupName -SnapshotName $snapshotName;;

    Write-Output "Creating disk $diskName"
    $disConfig = New-AzDiskConfig -Location $snapshot.Location -SourceResourceId $snapshot.Id -CreateOption Copy;
    $disk = New-AzDisk -Disk $disConfig -ResourceGroupName $testResourceGroupName -DiskName $diskName;

    Write-Output "Setting the vm size to $vmSize";
    $vm = New-AzVMConfig -VMName $vmName -VMSize $vmSize;
    $vm = Set-AzVMOSDisk -VM $vm -ManagedDiskId $disk.Id -CreateOption Attach -Windows;

    Write-Output "Creating network for $vmName";
    $publicIpName = ($vmName.ToLower()+'_ip');
    $publicIp = New-AzPublicIpAddress -Name $publicIpName -ResourceGroupName $testResourceGroupName -Location $snapshot.Location -AllocationMethod Dynamic;
    $privateIpName = ($vmName.ToLower()+'_pip');
    New-AzNetworkInterfaceIpConfig -Name $privateIpName -Subnet $vnet.Subnets[0] -PrivateIpAddress $ipaddress -Primary;
    $nicName = ($vmName.ToLower()+'_nic');
    $nic = New-AzNetworkInterface -Name $nicName -ResourceGroupName $testResourceGroupName -Location $snapshot.Location -SubnetId $vnet.Subnets[0].Id -IpConfigurationName $privateIpName -PublicIpAddressId $publicIp.Id;
    $vm = Add-AzVMNetworkInterface -VM $vm -Id $nic.Id;

    Write-Output "Creating VM $vmName";
    New-AzVM -VM $vm -ResourceGroupName $testResourceGroupName -Location $snapshot.Location;
}



New-VM -vmName "dc" -snapshotName $dcSnapshotName -ipAddress "172.31.11.5";
New-VM -vmName "qam" -snapshotName $qamSnapshotName -ipAddress "172.31.11.4";


Write-Host "Start executing the startup script"
$fileUri = @("https://automationadmin.blob.core.windows.net/startup/startup.ps1")
$settings = @{"fileUris" = $fileUri};
$command = "powershell -ExecutionPolicy Unrestricted -File startup.ps1 ";
$command += Get-ExecutionCommand -Name "StorageAccountName" -Value $storageAccountName
$command += Get-ExecutionCommand -Name "StorageKey" -Value $storageKey
$command += Get-ExecutionCommand -Name "StorageConnection" -Value $storageConnection
$command += Get-ExecutionCommand -Name "OutlookVersion" -Value $outlookVersion
$command += Get-ExecutionCommand -Name "QamInstallerVersion" -Value $qamInstallerVersion
$command += Get-ExecutionCommand -Name "Branch" -Value $branch
$command += Get-ExecutionCommand -Name "ArtUserName" -Value $artUserName
$command += Get-ExecutionCommand -Name "ArtPassword" -Value $artPassword
$command += Get-ExecutionCommand -Name "TestComponent" -Value $testComponent
$command += Get-ExecutionCommand -Name "TestTags" -Value $testTags
$command += Get-ExecutionCommand -Name "InstallFeatures" -Value $installFeatures
$command += Get-ExecutionCommand -Name "AzAccount" -Value $azAccount
$command += Get-ExecutionCommand -Name "AzPassword" -Value $azPassword

$protectedSettings = @{"storageAccountName" = $storageAccountName; "storageAccountKey" = $storageKey; "commandToExecute" = $command};
#run command
$location = "WestUS";
$extensionName = "PrepareQAM";
Write-Host "Removing the extension"
#Remove-AzVmExtension -Name $extensionName  -VMName "qam" -ResourceGroupName $testResourceGroupName -Force

Write-Host "Executing the extension Command: $command"
Set-AzVMExtension -ResourceGroupName $testResourceGroupName `
    -Location $location `
    -VMName "qam" `
    -Name $extensionName `
    -Publisher "Microsoft.Compute" `
    -ExtensionType "CustomScriptExtension" `
    -TypeHandlerVersion "1.9" `
    -Settings $settings    `
    -ProtectedSettings $protectedSettings