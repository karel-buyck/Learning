
# Eerst stoppen en dan deallocaten


Add-AzureRmAccount


# Set variables
     $resourceGroup = "RG-PLAYBIZ-NAV"
     $oldvmName = "PLAYSQL"
     $newvmName = "PLAYSQL2"


# Get the details of the VM to be renamed
     $originalVM = Get-AzureRmVM `
        -ResourceGroupName $resourceGroup `
        -Name $oldvmName


# Remove the original VM
     Remove-AzureRmVM -ResourceGroupName $resourceGroup -Name $oldvmName   


# Create the basic configuration for the replacement VM
     $newVM = New-AzureRmVMConfig -VMName $newvmName -VMSize $originalVM.HardwareProfile.VmSize


    Set-AzureRmVMOSDisk -VM $newVM -CreateOption Attach -ManagedDiskId $originalVM.StorageProfile.OsDisk.ManagedDisk.Id -Name $originalVM.StorageProfile.OsDisk.Name -Windows


# Add Data Disks
     foreach ($disk in $originalVM.StorageProfile.DataDisks) {
     Add-AzureRmVMDataDisk -VM $newVM `
        -Name $disk.Name `
        -ManagedDiskId $disk.ManagedDisk.Id `
        -Caching $disk.Caching `
        -Lun $disk.Lun `
        -DiskSizeInGB $disk.DiskSizeGB `
        -CreateOption Attach
     }


# Add NIC(s)
     foreach ($nic in $originalVM.NetworkProfile.NetworkInterfaces) {
         Add-AzureRmVMNetworkInterface `
            -VM $newVM `
            -Id $nic.Id
     }


# Recreate the VM
     New-AzureRmVM `
        -ResourceGroupName $resourceGroup `
        -Location $originalVM.Location `
        -VM $newVM `
        -DisableBginfoExtension