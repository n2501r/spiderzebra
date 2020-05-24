#############################################################################################
### Functions
#############################################################################################

###############################################
Function Function_One {
    write-host "Function One executed!"
}
###############################################
Function Function_Two {
    write-host "Function Two executed!"
}
###############################################
Function Function_Three {
    write-host "Function Three executed!"
}
###############################################
Function Function_Four {
    write-host "Function Four executed!"
}
###############################################
Function Function_Five {
    write-host "Function Five executed!"
}
###############################################


#############################################################################################
### Menu
#############################################################################################

Write-Host "~~~~~~~~~~~~~~~~~~ Menu Title ~~~~~~~~~~~~~~~~~~" -ForegroundColor Cyan
Write-Host "1: Enter 1 to execute Function One"
Write-Host "2: Enter 2 to execute Function Two"
Write-Host "3: Enter 3 to execute Function Three"
Write-Host "4: Enter 4 to execute Function Four"
Write-Host "5: Enter 5 to execute Function Five"
Write-Host "Q: Enter Q to quit."

$input = (Read-Host "Please make a selection").ToUpper()

switch ($input)
{
    '1' {Function_One}    ### Input the name of the function you want to execute when 1 is entered
    '2' {Function_Two}    ### Input the name of the function you want to execute when 2 is entered
    '3' {Function_Three}  ### Input the name of the function you want to execute when 3 is entered
    '4' {Function_Four}   ### Input the name of the function you want to execute when 4 is entered
    '5' {Function_Five}   ### Input the name of the function you want to execute when 5 is entered
    'Q' {Write-Host "The script has been canceled" -BackgroundColor Red -ForegroundColor White}
    Default {Write-Host "Your selection = $input, is not valid. Please try again." -BackgroundColor Red -ForegroundColor White}
}




#############################################################################################
### Cleanup
#############################################################################################

### Utilize the below code if you are running from powershell
Write-Host "Press any key to exit..."
$Readkey = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
