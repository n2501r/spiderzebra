#############################################################################################
### Out-GridView example with input directly from command.  This example will just display the services to the user.
#############################################################################################
Get-Service | Out-GridView -Title "List of all the services on this system" -OutputMode None




#############################################################################################
### Out-GridView example with input from variable.  This example will prompt the user to select one service they want to stop.
#############################################################################################
$services = Get-Service ### Getting all services on the system and sending them to the services variable
$single_output = $services | Out-GridView -Title "Select the service that you want to stop" -OutputMode Single ### Sending the services variable to Out-GridView, the single service selected will be sent to single_output variable

Write-Host "Stopping service" $single_output.Name -BackgroundColor Cyan -ForegroundColor Black
Stop-Service $single_output.Name ### Finally the service selected will be stopped




#############################################################################################
### Out-GridView example with input from custom array.  This example will prompt the user to select multiple running services they want to stop.
#############################################################################################
$custom_array = @() ### Creating an empty array to populate data in
$services = Get-Service | Select Name, Status, DisplayName, StartType | Where-Object {$_.Status -eq "Running"} ### Getting Running services on the system and sending them to the services variable

### Looping through all of the running services and adding them to the custom array
Foreach ($service in $services){ 

    ### Setting up custom array utilizing a PSObject
    $custom_array += New-Object PSObject -Property @{
    Service_Name = $service.Name
    Service_Status = $service.Status
    Service_DisplayName = $service.DisplayName
    Service_StartType = $service.StartType
    }
}

$multiple_output = $custom_array | Out-GridView -Title "This is using a custom array with multiple output values" -OutputMode Multiple

### Looping through all of the selected services and stopping the service
Foreach ($output in $multiple_output){

    Write-Host "Stopping service" $output.Service_Name -BackgroundColor Cyan -ForegroundColor Black
    Stop-Service $output.Service_Name ### Finally the service selected will be stopped

}
