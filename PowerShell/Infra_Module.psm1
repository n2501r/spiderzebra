#####################################################################################################################
### Function to randomly grab one infrastructure server based on the given site, environment, type and active status
#####################################################################################################################
Function Infra_One_Random ($Site, $Environment, $Type, $Active) { 
    ### SQL query used for invoke-sqlcmd
    $Infra_query="
    SELECT [site]
          ,[location]
          ,[server]
          ,[type]
          ,[active]
          ,[environment]
    FROM [demo_db].[dbo].[demo_table]
    GO
    "
    #####################################################
    try{
        ### Execution of the SQL Query and then it pipes the results to the where-object command to filter out the results based on the given parameters
        $Servers_Selected = Invoke-Sqlcmd -ServerInstance SQL_01 -Query "$Infra_query" -ErrorAction SilentlyContinue | Where-Object {$_.site -eq $Site -and $_.environment -eq $Environment -and $_.type -eq $Type -and $_.active -eq $Active}
        ### Randomly selecting one server from the returned servers from above query
        Get-Random -InputObject $Servers_Selected.server -Count 1
    }
    Catch [Exception]{ ### If the command inside the try statement fails the error will be outputted
        $Error_Message = "FAILED - "  ### Adding FAILED - before the actual error message to keep all the error messages in a standard format
        Return $Error_Message + $($_.Exception.Message)
    }
    #####################################################
}

#####################################################################################################################
### Function to grab all infrastructure servers based on the given site, environment, type and active status
#####################################################################################################################
Function Infra_All_Servers ($Site, $Environment, $Type, $Active) {
    ### SQL query used for invoke-sqlcmd
    $Infra_query="
    SELECT [site]
          ,[location]
          ,[server]
          ,[type]
          ,[active]
          ,[environment]
    FROM [demo_db].[dbo].[demo_table]
    GO
    "
    #####################################################
    try{
        ### Execution of the SQL Query and then it pipes the results to the where-object command to filter out the results based on the given parameters, then it pipes that to foreach-object to only return the name of the servers
        Invoke-Sqlcmd -ServerInstance SQL_01 -Query "$Infra_query" -ErrorAction SilentlyContinue | Where-Object {$_.type -eq $Type -and $_.site -eq $Site -and $_.active -eq $Active -and $_.environment -eq $Environment} | ForEach-Object {$_.server}
    }
    Catch [Exception]{ ### If the command inside the try statement fails the error will be outputted
        $Error_Message = "FAILED - " ### Adding FAILED - before the actual error message to keep all the error messages in a standard format
        Return $Error_Message + $($_.Exception.Message)
    }
    #####################################################
}

#####################################################################################################################
### Function to grab one random infrastructure server from each site based on the given environment, type and active status
#####################################################################################################################
Function Infra_One_Random_Each_Site ($Environment, $Type, $Active) {
    $Unique_Active_Servers = @() ### Creating empty array to be used later
    ### SQL query used for invoke-sqlcmd
    $Infra_query="
    SELECT [site]
          ,[location]
          ,[server]
          ,[type]
          ,[active]
          ,[environment]
    FROM [demo_db].[dbo].[demo_table]
    GO
    "
    #####################################################
    try{
        ### Execution of the SQL Query and then it pipes the results to the where-object command to filter out the results based on the given parameters
        $Servers_Selected = Invoke-Sqlcmd -ServerInstance SQL_01 -Query "$Infra_query" -ErrorAction SilentlyContinue | Where-Object {$_.active -eq $Active -and $_.type -eq $Type -and $_.environment -eq $Environment}
    }
    Catch [Exception]{ ### If the command inside the try statement fails the error will be outputted
        $Error_Message = "FAILED - " ### Adding FAILED - before the actual error message to keep all the error messages in a standard format
        Return $Error_Message + $($_.Exception.Message)
    }
    #####################################################
    ### Sorting the SQL data for just all of the unique sites so we can identify which server to use for each site
    $Unique_Sites = $Servers_Selected | Sort-Object -Unique -Property site | ForEach-Object{$_.site}
    ### Going through each unique site to identify one random server to use
    Foreach ($Site in $Unique_Sites){
        $Unique_Active_Servers += $Servers_Selected | Where-Object {$_.site -eq "$Site"} | Get-Random -Count 1 | ForEach-Object {$_.server}
    }
    Return $Unique_Active_Servers
}
