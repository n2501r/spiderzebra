###############################################################################
### Variables
###############################################################################
$Known_Hosts_File = $PSScriptRoot + "\NetNeighbors.csv" ### Location for the known hosts csv file, located in the same directory as this script
$Time = Get-Date -Format "MM-dd-yyyy-HHmm" ### Getting the date and formatting it to be used in the transcript file name
$Transcript_File = $PSScriptRoot + "\NetNeighbor_Report_$Time.txt" ### Location for the transcript file, located in the same directory as this script
$From = "NetNeighbor Watch <demo@gmail.com>" ### The email address to use to send the emails
$To = "demo@gmail.com" ### The email address to receive the emails
$SMTPServer = "smtp.gmail.com" ### smtp server
$SMTPPort = "587" ### smtp sever port number
$User = "demo" ### user id of the user to send the emails
$Password  = convertto-securestring "demopassword" -asplaintext -force ### converting plain text password to secure string
$creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $User, $Password ### putting together the PSCredential

### Cleaning up the old transcripts, only keeping the last 10 files which is 10 weeks
Get-ChildItem $PSScriptRoot -Recurse -Include NetNeighbor_Report*.txt| Where-Object {-not $_.PsIsContainer} | Sort-Object LastWriteTime -desc | Select-Object -Skip 10 | Remove-Item -Force

### Starting the transcript
Start-Transcript -Path $Transcript_File -Force
###############################################################################
### Function to send email with an attachment and body as html
###############################################################################
Function Send_Mail_HTML ($Body, $Subject, $Attachment) {
    Write-Host "[Sending Email] ($Subject) [$BODY]"
    Send-MailMessage -Attachments $Attachment -BodyAsHtml -From $From -to $To -Subject $Subject -Body $Body -SmtpServer $SMTPServer -port $SMTPPort -UseSsl -Credential $creds -WarningAction Ignore
}

###############################################################################
### Function to send standard email
###############################################################################
Function Send_Mail ($Body, $Subject) {
    Write-Host "[Sending Email] ($Subject) [$BODY]"
    Send-MailMessage -From $From -to $To -Subject $Subject -Body $Body -SmtpServer $SMTPServer -port $SMTPPort -UseSsl -Credential $creds -WarningAction Ignore
}
###############################################################################
### Function to build the known hosts report and send out via email
###############################################################################
Function Known_Hosts_Report {
    Write-Host "[Importing known hosts csv file]"
    ### Importing the known hosts csv file to be used for the report
    $Known_Hosts = Import-Csv -Path $Known_Hosts_File | Sort-Object { $_.IPAddress -as [Version]}
    ### Measuring how many hosts are present in the csv file
    $Known_Hosts_Count = $Known_Hosts | Measure-Object | ForEach-Object {$_.Count}
    Write-Host "[$Known_Hosts_Count host(s) in known hosts file]"
    #############################################################################################
    ### Building the HTML header and table column titles
    #############################################################################################
    $html_header = "
    <html>
    <head>
    <meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>
    <title>NetNeighbor Watch Report</title>
    <STYLE TYPE='text/css'>
    </style>
    </head>
    <body>
    <table-layout: fixed>
    <table width='100%'>
    <tr bgcolor='#4682B4'>
    <td colspan='7' height='25' align='center'><strong><font color='white' size='4' face='tahoma'>NetNeighbor Watch Report: $Known_Hosts_Count host(s)</font>
    </tr>
    </table>
    <table width='100%'>
    <tr bgcolor='#CCCCCC'>
    <td colspan='7' height='20' align='center'><strong><font color='black' size='2' face='tahoma'> Report of all known hosts that NetNeighbor Watch has discovered </font>
    </tr>
    </table>
    <table width='100%'><tbody>
        <tr bgcolor=black>
        <td width='25%' height='15' align='center'> <strong> <font color='white' size='2' face='tahoma' >IP Address</font></strong></td>
        <td width='25%' height='15' align='center'> <strong> <font color='white' size='2' face='tahoma' >MAC Address</font></strong></td>
        <td width='25%' height='15' align='center'> <strong> <font color='white' size='2' face='tahoma' >DNS Name</font></strong></td>
        <td width='25%' height='15' align='center'> <strong> <font color='white' size='2' face='tahoma' >Date Discovered</font></strong></td>
        </tr>
    </table>
    "

    #############################################################################################
    ### Going through each known host and attempting to resolve the dns name, then appending to the html_table variable
    #############################################################################################
    Foreach ($Known_Host in $Known_Hosts){
        Write-Host "[Processing: $($Known_Host.IPAddress)]"
        ### Attempting to resolve the IP Address to Name 
        $DNS_Name = Resolve-DnsName $($Known_Host.IPAddress) -QuickTimeout -ErrorAction SilentlyContinue | ForEach-Object {$_.NameHost}
        ### If there is no DNS name for that IP address it will set the variable to "DNS record does not exit"
        If ($null -eq $DNS_Name){$DNS_Name = "DNS record does not exist"}
        ### If htmlwrite_count variable is an even number the html background color for that row will be off-white, odd number will be gray
        $htmlwrite_count | ForEach-Object {if($_ % 2 -eq 0 ) {$htmlbgcolor = '<tr bgcolor=#F5F5F5>'} } ## Even Number (off-white)
        $htmlwrite_count | ForEach-Object {if($_ % 2 -eq 1 ) {$htmlbgcolor = '<tr bgcolor=#CCCCCC>'} } ## Odd Number (gray)
        #### Creating the HTML rows
        $html_table += "
        <table width='100%'><tbody>
            $htmlbgcolor
            <td width='25%' align='center'>$($Known_Host.IPAddress)</td>
            <td width='25%' align='center'>$($Known_Host.LinkLayerAddress)</td>
            <td width='25%' align='center'>$DNS_Name</td>
            <td width='25%' align='center'>$($Known_Host.Date_Discovered)</td>
            </tr>
        </table>
        "
        $htmlwrite_count++ ### Incrementing the count by 1 so that the next HTML row is a different color
    }
    #############################################################################################
    ### Gathering the latest transcript log from NetNeighbor_Watch.ps1 and sending email with the html report
    #############################################################################################
    ### Selecting the last modified transcript file
    Write-Host "[Gathering last modified transcript]"
    $Recent_Log = Get-ChildItem $PSScriptRoot -Recurse -Include NetNeighbor_Watch*.txt | Where-Object {-not $_.PsIsContainer} | Sort-Object LastWriteTime | Select-Object -Last 1
    ### Reading the content of the transcript file and sending it to NetNeighbor_Transcript.txt
    Get-Content -Path $Recent_Log -Raw | Out-File -FilePath $PSScriptRoot\NetNeighbor_Transcript.txt -Force
    ### Putting the html peices together
    $html_email = $html_header + $html_table
    ### Sending email with html report and last transcript attached
    Send_Mail_HTML -Body $html_email -Subject "[NetNeighbor Watch] Report: $Known_Hosts_Count host(s)" -Attachment $PSScriptRoot\NetNeighbor_Transcript.txt
}

###############################################################################
### Running the function(s)
###############################################################################
try{
    Known_Hosts_Report
}
Catch [Exception]{ ### If the command inside the try statement fails the error will be outputted
    $errormessage = $_.Exception.Message
    Write-host $errormessage -BackgroundColor Red -ForegroundColor White
    Send_Mail -Body "$errormessage" -Subject "[NetNeighbor Watch] Error Triggered"
}

Stop-Transcript
