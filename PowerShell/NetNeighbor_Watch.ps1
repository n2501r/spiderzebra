###############################################################################
### Variables
###############################################################################
$subnet = "10.0.0" ### Subnet that you want to scan for new hosts
$Gateway = Get-NetRoute -DestinationPrefix 0.0.0.0/0 | Select-Object -ExpandProperty NextHop ### Default gateway IP address pulled from the Default Route
$start = 0 ### The starting number for the last octect of the subnet you want to scan for new hosts
$end = 255 ### The ending number for the last octect of the subnet you want to scan for new hosts
$Known_Hosts_File = $PSScriptRoot + "\NetNeighbors.csv" ### Location for the known hosts csv file, located in the same directory as this script
$Time = Get-Date -Format "MM-dd-yyyy-HHmm" ### Getting the date and formatting it to be used in the transcript file name
$Transcript_File = $PSScriptRoot + "\NetNeighbor_Watch_$Time.txt" ### Location for the transcript file, located in the same directory as this script
$From = "NetNeighbor Watch <demo@gmail.com>" ### The email address to use to send the emails
$To = "demo@gmail.com" ### The email address to receive the emails
$SMTPServer = "smtp.gmail.com" ### smtp server
$SMTPPort = "587" ### smtp sever port number
$User = "demo" ### user id of the user to send the emails
$Password  = convertto-securestring "demopassword" -asplaintext -force ### converting plain text password to secure string
$creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $User, $Password ### putting together the PSCredential

### Cleaning up the old transcripts, only keeping the last 300 files which is about 7 days
Get-ChildItem $PSScriptRoot -Recurse -Include NetNeighbor_Watch*.txt | Where-Object {-not $_.PsIsContainer} | Sort-Object LastWriteTime -desc | Select-Object -Skip 300 | Remove-Item -Force

### Starting the transcript
Start-Transcript -Path $Transcript_File -Force
###############################################################################
### Function to send email
###############################################################################
Function Send_Mail ($Body, $Subject) {
    Write-Host "[Sending Email] ($Subject) [$BODY]"
    Send-MailMessage -From $From -to $To -Subject $Subject -Body $Body -SmtpServer $SMTPServer -port $SMTPPort -UseSsl -Credential $creds -WarningAction Ignore
}

###############################################################################
### Function to do a test-connection on specified subnet to identify new hosts
###############################################################################
Function Ping_Sweep {
    Write-Host "[Ping Sweep]"
    while ($start -le $end) { ### Conduct the while loop until the start variable is greater than the end variable
        $IP = "$subnet.$start" ### Combining the subnet variable with the start variable to create the IP address to ping
        $Connected = Test-Connection -ComputerName $IP -count 1 -Quiet -Delay 1 -TimeoutSeconds 1 ### Pinging the given IP address
        If ($Connected -eq $True){ ### If the IP is pingable then it will output that the IP is up
            Write-Host "[UP] $IP"
        }
        $start++ ### Incrementing the start variable by 1 each time through the loop
    }
}
###############################################################################
### Function to discover all neighbors
###############################################################################
Function NetNeighbors {
    Write-Host "[Identifying Net Neighbors]"
    ### Pulling Address Resolution Protocol (ARP) cache for only IPv4 addresses in a stale or reachable state and for only IP addresses that are in the given subnet
    $Neighbors = Get-NetNeighbor -AddressFamily IPv4 -State Stale, Reachable | Select-Object IPAddress, LinkLayerAddress | Where-Object {$_.IPAddress -match $subnet}
    Return $Neighbors
}

###############################################################################
### Function to import the known list of hosts
###############################################################################
Function Known_Hosts {
    Write-Host "[Importing Known Hosts]"
    ### Importing the known hosts csv file to be used for comparing against the ARP cache
    $Known_Hosts = Import-Csv -Path $Known_Hosts_File 
    Return $Known_Hosts
}

###############################################################################
### Function to compare the results of the net neighbors and the import of the known hosts to determine any new hosts
###############################################################################
Function Compare_Results ($NetNeighbors, $Known_Hosts) {
    Write-Host "[Comparing Net Neighbors and Known Hosts for differences]"
    ### Comparing the known hosts to the ARP cache for any newly discovered hosts
    $Results = Compare-Object -ReferenceObject $Known_Hosts -DifferenceObject $NetNeighbors -Property IPAddress, LinkLayerAddress 
    If ($Null -eq $Results){Write-Host "[No Change] Net Neighbors and Known Hosts Match, nothing to do here..."}
    Else{ ### If there is a new host discovered it will be processed
        Foreach ($Result in $Results){
            If ($Result.SideIndicator -eq '=>'){ ### If the sideindicator equals => it means that a new host was present in the NetNeighbors
                ### Creating a PSCustomObject with the IPAddress, LinkLayerAddress and Date
                $outputcsv = [PSCustomObject] @{
                IPAddress = $Result.IPAddress
                LinkLayerAddress = $Result.LinkLayerAddress
                Date_Discovered = (Get-Date)
                }
                ### Appending the new host information to the known hosts file so that we won't get alerted on this host again
                $outputcsv | Export-Csv -Path $Known_Hosts_File -Append -NoClobber
                ### Attempting to resolve the IP Address to Name to include in the email alert
                $DNS_Name = Resolve-DnsName $($Result.IPAddress) -QuickTimeout -ErrorAction SilentlyContinue | Select-Object -ExpandProperty NameHost
                ### If there is no DNS name for that IP address it will set the variable to "DNS record does not exit"
                If ($null -eq $DNS_Name){$DNS_Name = "DNS record does not exist"}
                ### Calling the Send_Mail function with the new host specific information
                Send_Mail -Body "$($Result.IPAddress) / $($Result.LinkLayerAddress) / $DNS_Name" -Subject "[NetNeighbor Watch] Alert: New Host Detected"
            }
            If ($Result.SideIndicator -eq '<='){ ### If the sideindicator equals <= it means that a host was present in the known host csv file but not in ARP cache
                Write-Host "[Offline Host] $($Result.IPAddress) / $($Result.LinkLayerAddress) is present in the known hosts csv file but not in ARP cache"
            }
        }
    }
}

###############################################################################
### Function to conduct a pre check to ensure there is a good csv file present
###############################################################################
Function Pre_Check {
    If (-not (Test-Path $Known_Hosts_File)) {  ### if there is not a known host csv file, the script will create one
        Write-Host "[Pre Check] No csv file present, creating one..."
        ### Grabbing the LinkLayerrAddress of the default gateway so that it can populate the known host csv file
        $Gateway = Get-NetNeighbor -AddressFamily IPv4 -State Stale, Reachable | Select-Object IPAddress, LinkLayerAddress | Where-Object {$_.IPAddress -eq $Gateway}
        ### Creating a PSCustomObject with the IPAddress, LinkLayerAddress and Date
        $outputcsv = [PSCustomObject] @{
        IPAddress = $Gateway.IPAddress
        LinkLayerAddress = $Gateway.LinkLayerAddress
        Date_Discovered = (Get-Date)
        }
        ### writing the default gateway information to the known hosts file
        $outputcsv | Export-Csv -Path $Known_Hosts_File
    }
    Else{ ### If there is a known host csv file, the script will see if its blank
        $Import_Check = Import-Csv -Path $Known_Hosts_File ### importing the known host csv file
        If ($null -eq $Import_Check){ ### if the known host csv file is blank it will populat it with the default gateway
            Write-host "[Empty CSV File, script will attempt to populate it with the gateway]" -BackgroundColor Red -ForegroundColor White
            Send_Mail -Body "Empty CSV File, script will attempt to populate it with the gateway." -Subject "[NetNeighbor Watch] Error Triggered"
            ### Grabbing the LinkLayerrAddress of the default gateway so that it can populate the known host csv file
            $Gateway = Get-NetNeighbor -AddressFamily IPv4 -State Stale, Reachable | Select-Object IPAddress, LinkLayerAddress | Where-Object {$_.IPAddress -eq $Gateway}
            ### Creating a PSCustomObject with the IPAddress, LinkLayerAddress and Date   
            $outputcsv = [PSCustomObject] @{
            IPAddress = $Gateway.IPAddress
            LinkLayerAddress = $Gateway.LinkLayerAddress
            Date_Discovered = (Get-Date)
            }
            ### writing the default gateway information to the known hosts file
            $outputcsv | Export-Csv -Path $Known_Hosts_File -Force
        }
    }
}

###############################################################################
### Running the functions
###############################################################################
try{
    Pre_Check
    Ping_Sweep
    $NetNeighbors = NetNeighbors
    $Known_Hosts = Known_Hosts
    Compare_Results -NetNeighbors $NetNeighbors -Known_Hosts $Known_Hosts 
}
Catch [Exception]{ ### If the command inside the try statement fails the error will be outputted
    $errormessage = $_.Exception.Message
    Write-host $errormessage -BackgroundColor Red -ForegroundColor White
    Send_Mail -Body "$errormessage" -Subject "[NetNeighbor Watch] Error Triggered"
}

Stop-Transcript
