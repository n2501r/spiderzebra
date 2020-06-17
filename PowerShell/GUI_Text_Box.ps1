###############################################################################
### Function to provide a graphical input box to accept input from
###############################################################################
Function GUI_TextBox ($Input_Type){

    ### Creating the form with the Windows forms namespace
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Enter the appropriate information' ### Text to be displayed in the title
    $form.Size = New-Object System.Drawing.Size(310,625) ### Size of the window
    $form.StartPosition = 'CenterScreen'  ### Optional - specifies where the window should start
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedToolWindow  ### Optional - prevents resize of the window
    $form.Topmost = $true  ### Optional - Opens on top of other windows

    ### Adding an OK button to the text box window
    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = New-Object System.Drawing.Point(155,550) ### Location of where the button will be
    $OKButton.Size = New-Object System.Drawing.Size(75,23) ### Size of the button
    $OKButton.Text = 'OK' ### Text inside the button
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $OKButton
    $form.Controls.Add($OKButton)

    ### Adding a Cancel button to the text box window
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Point(70,550) ### Location of where the button will be
    $CancelButton.Size = New-Object System.Drawing.Size(75,23) ### Size of the button
    $CancelButton.Text = 'Cancel' ### Text inside the button
    $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $CancelButton
    $form.Controls.Add($CancelButton)

    ### Putting a label above the text box
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10,10) ### Location of where the label will be
    $label.AutoSize = $True
    $Font = New-Object System.Drawing.Font("Arial",12,[System.Drawing.FontStyle]::Bold) ### Formatting text for the label
    $label.Font = $Font
    $label.Text = $Input_Type ### Text of label, variable used to provide more information to the user
    $label.ForeColor = 'Red' ### Color of the label text
    $form.Controls.Add($label)

    ### Inserting the text box that will accept input
    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point(10,40) ### Location of the text box
    $textBox.Size = New-Object System.Drawing.Size(275,500) ### Size of the text box
    $textBox.Multiline = $true ### Allows multiple lines of data
    $textbox.AcceptsReturn = $true ### By hitting enter it creates a new line
    $textBox.ScrollBars = "Vertical" ### Allows for a vertical scroll bar if the list of text is too big for the window
    $form.Controls.Add($textBox)

    $form.Add_Shown({$textBox.Select()}) ### Activates the form and sets the focus on it
    $result = $form.ShowDialog() ### Displays the form 

    ### If the OK button is selected do the following
    if ($result -eq [System.Windows.Forms.DialogResult]::OK)
    {
        ### Removing all the spaces and extra lines
        $x = $textBox.Lines | Where{$_} | ForEach{ $_.Trim() }
        ### Putting the array together
        $array = @()
        ### Putting each entry into array as individual objects
        $array = $x -split "`r`n"
        ### Sending back the results while taking out empty objects
        Return $array | Where-Object {$_ -ne ''}
    }

    ### If the cancel button is selected do the following
    if ($result -eq [System.Windows.Forms.DialogResult]::Cancel)
    {
        Write-Host "User Canceled" -BackgroundColor Red -ForegroundColor White
        Write-Host "Press any key to exit..."
        $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        Exit
    }

}
###############################################################################



###############################################################################
### Computer Name(s) example of how to utilize the GUI_TextBox function
###############################################################################
$Computers = GUI_TextBox "Computer Names(s):" ### Calls the text box function with a parameter and puts returned input in variable
$Computer_Count = $Computers | Measure-Object | % {$_.Count} ### Measures how many objects were inputted

If ($Computer_Count -eq 0){ ### If the count returns 0 it will throw and error
    Write-Host "Nothing was inputed..." -BackgroundColor Red -ForegroundColor White
    Return
}
Else { ### If there was actual data returned in the input, the script will continue
    Write-Host "Number of computers entered:" $Computer_Count -BackgroundColor Cyan -ForegroundColor Black
    $Computers
    ### Here is where you would put your specific code to take action on those computers inputted
}

###############################################################################
### User Name(s) example of how to utilize the GUI_TextBox function
###############################################################################
$Users = GUI_TextBox "User Names(s):" ### Calls the text box function with a parameter and puts returned input in variable
$User_Count = $Users | Measure-Object | % {$_.Count} ### Measures how many objects were inputted

If ($User_Count -eq 0){ ### If the count returns 0 it will throw and error
    Write-Host "Nothing was inputed..." -BackgroundColor Red -ForegroundColor White
    Return
}
Else { ### If there was actual data returned in the input, the script will continue
    Write-Host "Number of users entered:" $User_Count -BackgroundColor Cyan -ForegroundColor Black
    $Users
    ### Here is where you would put your specific code to take action on those users inputted
}
