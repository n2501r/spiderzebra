#############################################################################################
### Variables
#############################################################################################
$htmlfile = "\\server\full\path\to\file.html" ### The full path to the final HTML file
$htmlfile_temp = "\\server\full\path\to\file_temp.html" ### The full path to the temporary HTML file

### Checking to see if the temp file exists, if it does it will remove it
If (Test-Path $htmlfile_temp) { Remove-Item $htmlfile_temp -Force }

#############################################################################################
### Building the HTML header and table column titles
#############################################################################################
$html_header = "
<html>
<head>
<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>
<title>DHCP Report</title>
<STYLE TYPE='text/css'>
</style>
</head>
<body>
<table-layout: fixed>
<table width='100%'>
<tr bgcolor='#00B624'>
<td colspan='7' height='25' align='center'><strong><font color='#000000' size='4' face='tahoma'>DHCP Scope Statistics Report</font><font color='#000000' size='4' face='tahoma'> ($(Get-Date))</font><font color='#000000' size='2' face='tahoma'> <BR> Data Updates Every Day</font>
</tr>
</table>
<table width='100%'>
<tr bgcolor='#CCCCCC'>
<td colspan='7' height='20' align='center'><strong><font color='#000000' size='2' face='tahoma'><span style=background-color:#FFF284>WARNING</span> at 80% In Use &nbsp;&nbsp;&nbsp;&nbsp; <span style=background-color:#FF0000><font color=white>CRITICAL</font></span> at 95% In Use</font>
</tr>
</table>
<table width='100%'><tbody>
    <tr bgcolor=black>
    <td width='10%' height='15' align='center'> <strong> <font color='white' size='2' face='tahoma' >DHCP Server</font></strong></td>
    <td width='8%' height='15' align='center'> <strong> <font color='white' size='2' face='tahoma' >Scope ID</font></strong></td>
    <td width='10%' height='15' align='center'> <strong> <font color='white' size='2' face='tahoma' >Scope name</font></strong></td>
    <td width='8%' height='15' align='center'> <strong> <font color='white' size='2' face='tahoma' >Scope State</font></strong></td>
    <td width='8%' height='15' align='center'> <strong> <font color='white' size='2' face='tahoma' >In Use</font></strong></td>
    <td width='8%' height='15' align='center'> <strong> <font color='white' size='2' face='tahoma' >Free</font></strong></td>
    <td width='8%' height='15' align='center'> <strong> <font color='white' size='2' face='tahoma' >% In Use</font></strong></td>
    <td width='8%' height='15' align='center'> <strong> <font color='white' size='2' face='tahoma' >Reserved</font></strong></td>
    <td width='8%' height='15' align='center'> <strong> <font color='white' size='2' face='tahoma' >Subnet Mask</font></strong></td>
    <td width='8%' height='15' align='center'> <strong> <font color='white' size='2' face='tahoma' >Start of Range</font></strong></td>
    <td width='8%' height='15' align='center'> <strong> <font color='white' size='2' face='tahoma' >End of Range</font></strong></td>
    <td width='8%' height='15' align='center'> <strong> <font color='white' size='2' face='tahoma' >Lease Duration</font></strong></td>
    </tr>
</table>
"
$html_header | Out-File $htmlfile_temp ### Writing the HTML header to the temporary file

#############################################################################################
### DHCP Statistic Gathering
#############################################################################################
$DHCP_Servers = Get-DhcpServerInDC | ForEach-Object {$_.DnsName} | Sort-Object -Property DnsName ### Dynamically pulling the DHCP servers in a Active Directory domain
Foreach ($DHCP_Server in $DHCP_Servers){ ### Going through the DHCP servers that were returned one at a time to pull statistics
    $DHCP_Scopes = Get-DhcpServerv4Scope –ComputerName $DHCP_Server | Select-Object ScopeId, Name, SubnetMask, StartRange, EndRange, LeaseDuration, State ### Getting all the dhcp scopes for the given server
    Foreach ($DHCP_Scope in $DHCP_Scopes){ ### Going through the scopes returned in a given server
        $DHCP_Scope_Stats = Get-DhcpServerv4ScopeStatistics -ComputerName $DHCP_Server -ScopeId $DHCP_Scope.ScopeId | Select-Object Free, InUse, Reserved, PercentageInUse, ScopeId ### Gathering the scope stats
        $percentinuserounded = ([math]::Round($DHCP_Scope_Stats.PercentageInUse,0)) ### Rounding the percent in use to have no decimals
        ### Color formatting based on how much a scope is in use
        If ($percentinuserounded -ge 95){$htmlpercentinuse = '<td width="8%" align="center" td bgcolor="#FF0000"> <font color="white">' + $percentinuserounded + '</font></td>'}
        If ($percentinuserounded -ge 80 -and $percentinuserounded -lt 95){$htmlpercentinuse = '<td width="8%" align="center" td bgcolor="#FFF284"> <font color="black">' + $percentinuserounded + '</font></td>'}
        If ($percentinuserounded -lt 80){$htmlpercentinuse = '<td width="8%" align="center" td bgcolor="#A6CAA9"> <font color="black">' + $percentinuserounded + '</font></td>'}
        ### Changing the cell color if the scope is inactive / active
        If ($DHCP_Scope.State -eq "Inactive"){$htmlScopeState = '<td width="8%" align="center" td bgcolor="#AAAAB2"> <font color="black">' + $DHCP_Scope.State  + '</font></td>'}
        If ($DHCP_Scope.State -eq "Active"){$htmlScopeState = '<td width="8%" align="center">' + $DHCP_Scope.State + '</td>'}
        ### Changing the background color on every other scope so the html is easy to read
        $htmlwrite_count | ForEach-Object {if($_ % 2 -eq 0 ) {$htmlbgcolor = '<tr bgcolor=#F5F5F5>'} } ## Even Number (off-white)
        $htmlwrite_count | ForEach-Object {if($_ % 2 -eq 1 ) {$htmlbgcolor = '<tr bgcolor=#CCCCCC>'} } ## Odd Number (gray)
        #### Creating the HTML row for the given DHCP scope with the detailed stats and information
        $current = "
        <table width='100%'><tbody>
            $htmlbgcolor
            <td width='10%' align='center'>$($DHCP_Server.TrimEnd(".local.domain"))</td>
            <td width='8%' align='center'>$($DHCP_Scope.ScopeId)</td>
            <td width='10%' align='center'>$($DHCP_Scope.Name)</td>
            $htmlScopeState
            <td width='8%' align='center'>$($DHCP_Scope_Stats.InUse)</td>
            <td width='8%' align='center'>$($DHCP_Scope_Stats.Free)</td>
            $htmlpercentinuse
            <td width='8%' align='center'>$($DHCP_Scope_Stats.Reserved)</td>
            <td width='8%' align='center'>$($DHCP_Scope.SubnetMask)</td>
            <td width='8%' align='center'>$($DHCP_Scope.StartRange)</td>
            <td width='8%' align='center'>$($DHCP_Scope.EndRange)</td>
            <td width='8%' align='center'>$($DHCP_Scope.LeaseDuration)</td>
            </tr>
        </table>
        "
        $current  | Out-File $htmlfile_temp -Append ### Appending the HTML row to the tempory file

        $htmlwrite_count++ ### Incrementing the count by 1 so that the next HTML row is a different color
        Clear-Variable htmlScopeState, htmlpercentinuse, percentinuserounded, DHCP_Scope_Stats -ErrorAction SilentlyContinue
    }
}
Clear-Variable htmlwrite_count

#############################################################################################
### HTML file cleanup
#############################################################################################
If (Test-Path $htmlfile) { Remove-Item $htmlfile -Force } ### Removing the final html file if it exists
Rename-Item $htmlfile_temp $htmlfile -Force ### Renaming the temp file to the final file
