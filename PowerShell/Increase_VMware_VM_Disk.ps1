###############################################################################
### Variables
###############################################################################
### Launch GUI textbox to accept input from user
$VMs = GUI_TextBox "VM Name(s):" ### This function was introduced in previous blog post
$VM_Count = $VMs | Measure-Object | ForEach-Object {$_.Count}
If ($VM_Count -eq 0){ ### If nothing was inputed, the script will not continue
    Write-Host "Nothing was inputed, script is exiting..." -BackgroundColor Red -ForegroundColor White
    Return
}
Else{
    $Target_Disk = Read-Host "Enter the target disk size in GB (ie: 180)"
    $vCenter = Read-Host "Enter the vCenter FQDN you want to connect to"
}

###############################################################################
### Confirmation of the information that you entered
###############################################################################
Write-Host "Below are the variables that will be fed into the script:" -ForegroundColor Black -BackgroundColor Yellow
Write-Host "Number of VMs to increase disk:" $VM_Count -ForegroundColor White -BackgroundColor Blue
Write-Host "vCenter to connect to:" $vCenter -ForegroundColor White -BackgroundColor Blue
Write-Host "Target disk size in GB:" $Target_Disk -ForegroundColor White -BackgroundColor Blue
$final_confirm = (Read-Host "Do you want to Continue? (Y/N)").ToUpper()

###############################################################################
### Connect to VMware vCenter
###############################################################################
$VIcred = Get-Credential
Write-Host "Please wait while the script connects to vCenter: $vCenter" -BackgroundColor Cyan -ForegroundColor Black
Connect-VIServer -Server $vCenter -Credential $VIcred -ErrorAction Stop -Force

###############################################################################
### Execute the disk increase
###############################################################################
If ($final_confirm -eq "Y") { ### executes if the user entered y or Y on the final confiramtion question
    Foreach ($VM in $VMs){ ### going through each VM in the VMs variable
        Write-Host $VM "- Increasing disk to $Target_Disk GB" -BackgroundColor Yellow -ForegroundColor Black
        try{
            ### Resize the VM disk through VMware PowerCLI
            Get-HardDisk -Server $vCenter -VM $VM | Where-Object {$_.Name -eq "Hard Disk 1"} | Set-HardDisk -Server $vCenter -CapacityGB $Target_Disk -Confirm:$false
            ### Utilize invoke-command to expand the C partition within windows 10
            $ScriptBlock = { ### The commands that will be executed on the remote windows 10 VM
                $MaxSize = (Get-PartitionSupportedSize -DriveLetter C).Sizemax  ### This grabs the max size which should now be larger than the current partition
                Resize-Partition -DriveLetter C -Size $MaxSize ### command to expand the parition
            }
            Invoke-Command -ComputerName $VM -ScriptBlock $ScriptBlock ### execution of the scriptblock commands on the remote windows 10 VM
        }
        Catch [Exception]{ ### If the command inside the try statement fails the error will be outputted
            $errormessage = $_.Exception.Message
            Write-Host $errormessage -BackgroundColor Red -ForegroundColor White
        }
        Write-Host "$VM - Process Complete" -BackgroundColor Cyan -ForegroundColor Black
    }
}
Else{Write-Host "You selected Cancel or Nothing Present" -BackgroundColor Red -ForegroundColor White}

###############################################################################
### Disconnecting from the vCenter to keep things clean
###############################################################################
Disconnect-VIServer -Server $vCenter -Confirm:$false
