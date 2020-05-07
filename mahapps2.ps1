$global:getloaction = Get-Location

[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')  | out-null
[System.Reflection.Assembly]::LoadWithPartialName('System.ComponentModel') | out-null
[System.Reflection.Assembly]::LoadWithPartialName('System.Data')           | out-null
[System.Reflection.Assembly]::LoadWithPartialName('System.Drawing')        | out-null
[System.Reflection.Assembly]::LoadWithPartialName('presentationframework') | out-null

[System.Reflection.Assembly]::LoadWithPartialName('PresentationCore')      | out-null
[System.Reflection.Assembly]::LoadFrom("$($global:getloaction)\assembly\MahApps.Metro.dll")       | out-null
[System.Reflection.Assembly]::LoadFrom("$($global:getloaction)\assembly\System.Windows.Interactivity.dll") | out-null

Add-Type -AssemblyName "System.Windows.Forms"
Add-Type -AssemblyName "System.Drawing"

##I HIDE POWERSHELL WINDOWZ LEL##
##https://stackoverflow.com/questions/1802127/how-to-run-a-powershell-script-without-displaying-a-window##

$t = '[DllImport("user32.dll")] public static extern bool ShowWindow(int handle, int state);'
add-type -name win -member $t -namespace native
[native.win]::ShowWindow(([System.Diagnostics.Process]::GetCurrentProcess() | Get-Process).MainWindowHandle, 0)
Function Convert-FspToUsername 
{ 
    <# 
        .SYNOPSIS 
            Convert a FSP to a sAMAccountName 
        .DESCRIPTION 
            This function converts FSP's to sAMAccountName's. 
        .PARAMETER UserSID 
            This is the SID of the FSP in the form of S-1-5-20. These can be found 
            in the ForeignSecurityPrincipals container of your domain. 
        .EXAMPLE 
            Convert-FspToUsername -UserSID "S-1-5-11","S-1-5-17","S-1-5-20" 
 
            sAMAccountName                      Sid 
            --------------                      --- 
            NT AUTHORITY\Authenticated Users    S-1-5-11 
            NT AUTHORITY\IUSR                   S-1-5-17 
            NT AUTHORITY\NETWORK SERVICE        S-1-5-20 
 
            Description 
            =========== 
            This example shows passing in multipe sids to the function 
        .EXAMPLE 
            Get-ADObjects -ADSPath "LDAP://CN=ForeignSecurityPrincipals,DC=company,DC=com" -SearchFilter "(objectClass=foreignSecurityPrincipal)" | 
            foreach {$_.Properties.name} |Convert-FspToUsername 
 
            sAMAccountName                      Sid 
            --------------                      --- 
            NT AUTHORITY\Authenticated Users    S-1-5-11 
            NT AUTHORITY\IUSR                   S-1-5-17 
            NT AUTHORITY\NETWORK SERVICE        S-1-5-20 
 
            Description 
            =========== 
            This example takes the output of the Get-ADObjects function, and pipes it through foreach to get to the name 
            property, and the resulting output is piped through Convert-FspToUsername. 
        .NOTES 
            This function currently expects a SID in the same format as you see being displayed 
            as the name property of each object in the ForeignSecurityPrincipals container in your 
            domain.  
        .LINK 
            https://code.google.com/p/mod-posh/wiki/ActiveDirectoryManagement#Convert-FspToUsername 
    #> 
    [CmdletBinding()] 
    Param 
        ( 
        [Parameter(Position=0,Mandatory=$true,ValueFromPipeline=$true)] 
        $UserSID 
        ) 
    Begin 
    { 
        } 
    Process 
    { 
        foreach ($Sid in $UserSID) 
        { 
            try 
            { 
                $SAM = (New-Object System.Security.Principal.SecurityIdentifier($Sid)).Translate([System.Security.Principal.NTAccount]) 
                $Result = New-Object -TypeName PSObject -Property @{ 
                    Sid = $Sid 
                    sAMAccountName = $SAM.Value 
                    } 
                Return $Result 
                } 
            catch 
            { 
                $Result = New-Object -TypeName PSObject -Property @{ 
                    Sid = $Sid 
                    sAMAccountName = $Error[0].Exception.InnerException.Message.ToString().Trim() 
                    } 
                Return $Result 
                } 
            } 
        } 
    End 
    { 
        } 
    }
function Invoke-Sqlcmd2 {
    <#
        .SYNOPSIS
            Runs a T-SQL script.
        .DESCRIPTION
            Runs a T-SQL script. Invoke-Sqlcmd2 runs the whole script and only captures the first selected result set, such as the output of PRINT statements when -verbose parameter is specified.
            Parameterized queries are supported.
            Help details below borrowed from Invoke-Sqlcmd
        .PARAMETER ServerInstance
            Specifies the SQL Server instance(s) to execute the query against.
        .PARAMETER Database
            Specifies the name of the database to execute the query against. If specified, this database will be used in the ConnectionString when establishing the connection to SQL Server.
            If a SQLConnection is provided, the default database for that connection is overridden with this database.
        .PARAMETER Query
            Specifies one or more queries to be run. The queries can be Transact-SQL, XQuery statements, or sqlcmd commands. Multiple queries in a single batch may be separated by a semicolon.
            Do not specify the sqlcmd GO separator (or, use the ParseGo parameter). Escape any double quotation marks included in the string.
            Consider using bracketed identifiers such as [MyTable] instead of quoted identifiers such as "MyTable".
        .PARAMETER InputFile
            Specifies the full path to a file to be used as the query input to Invoke-Sqlcmd2. The file can contain Transact-SQL statements, XQuery statements, sqlcmd commands and scripting variables.
        .PARAMETER Credential
            Login to the target instance using alternative credentials. Windows and SQL Authentication supported. Accepts credential objects (Get-Credential)
            SECURITY NOTE: If you use the -Debug switch, the connectionstring including plain text password will be sent to the debug stream.
        .PARAMETER Encrypt
            If this switch is enabled, the connection to SQL Server will be made using SSL.
            This requires that the SQL Server has been set up to accept SSL requests. For information regarding setting up SSL on SQL Server, see https://technet.microsoft.com/en-us/library/ms189067(v=sql.105).aspx
        .PARAMETER QueryTimeout
            Specifies the number of seconds before the queries time out.
        .PARAMETER ConnectionTimeout
            Specifies the number of seconds before Invoke-Sqlcmd2 times out if it cannot successfully connect to an instance of the Database Engine. The timeout value must be an integer between 0 and 65534. If 0 is specified, connection attempts do not time out.
        .PARAMETER As
            Specifies output type. Valid options for this parameter are 'DataSet', 'DataTable', 'DataRow', 'PSObject', and 'SingleValue'
            PSObject output introduces overhead but adds flexibility for working with results: http://powershell.org/wp/forums/topic/dealing-with-dbnull/
        .PARAMETER SqlParameters
            Specifies a hashtable of parameters for parameterized SQL queries.  http://blog.codinghorror.com/give-me-parameterized-sql-or-give-me-death/
            Example:
        .PARAMETER AppendServerInstance
            If this switch is enabled, the SQL Server instance will be appended to PSObject and DataRow output.
        .PARAMETER ParseGo
            If this switch is enabled, "GO" statements will be handled automatically.
            Every "GO" will effectively run in a separate query, like if you issued multiple Invoke-SqlCmd2 commands.
            "GO"s will be recognized if they are on a single line, as this covers
            the 95% of the cases "GO" parsing is needed
            Note:
                Queries will always target that database, e.g. if you have this Query:
                    USE DATABASE [dbname]
                    GO
                    SELECT * from sys.tables
                and you call it via
                    Invoke-SqlCmd2 -ServerInstance instance -Database msdb -Query ...
                you'll get back tables from msdb, not dbname.
        .PARAMETER SQLConnection
            Specifies an existing SQLConnection object to use in connecting to SQL Server. If the connection is closed, an attempt will be made to open it.
        .PARAMETER ApplicationName
             If specified, adds the given string into the ConnectionString's Application Name property which is visible via SQL Server monitoring scripts/utilities to indicate where the query originated.
        .PARAMETER MessagesToOutput
            Use this switch to have on the output stream messages too (e.g. PRINT statements). Output will hold the resultset too. See examples for detail
            NB: only available from Powershell 3 onwards
        .INPUTS
            String[]
                You can only pipe strings to to Invoke-Sqlcmd2: they will be considered as passed -ServerInstance(s)
        .OUTPUTS
        As PSObject:     System.Management.Automation.PSCustomObject
        As DataRow:      System.Data.DataRow
        As DataTable:    System.Data.DataTable
        As DataSet:      System.Data.DataTableCollectionSystem.Data.DataSet
        As SingleValue:  Dependent on data type in first column.
        .EXAMPLE
            Invoke-Sqlcmd2 -ServerInstance "MyComputer\MyInstance" -Query "SELECT login_time AS 'StartTime' FROM sysprocesses WHERE spid = 1"
            Connects to a named instance of the Database Engine on a computer and runs a basic T-SQL query.
            StartTime
            -----------
            2010-08-12 21:21:03.593
        .EXAMPLE
            Invoke-Sqlcmd2 -ServerInstance "MyComputer\MyInstance" -InputFile "C:\MyFolder\tsqlscript.sql" | Out-File -filePath "C:\MyFolder\tsqlscript.rpt"
            Reads a file containing T-SQL statements, runs the file, and writes the output to another file.
        .EXAMPLE
            Invoke-Sqlcmd2  -ServerInstance "MyComputer\MyInstance" -Query "PRINT 'hello world'" -Verbose
            Uses the PowerShell -Verbose parameter to return the message output of the PRINT command.
            VERBOSE: hello world
        .EXAMPLE
            Invoke-Sqlcmd2 -ServerInstance MyServer\MyInstance -Query "SELECT ServerName, VCNumCPU FROM tblServerInfo" -as PSObject | ?{$_.VCNumCPU -gt 8}
            Invoke-Sqlcmd2 -ServerInstance MyServer\MyInstance -Query "SELECT ServerName, VCNumCPU FROM tblServerInfo" -as PSObject | ?{$_.VCNumCPU}
            This example uses the PSObject output type to allow more flexibility when working with results.
            If we used DataRow rather than PSObject, we would see the following behavior:
                Each row where VCNumCPU does not exist would produce an error in the first example
                Results would include rows where VCNumCPU has DBNull value in the second example
        .EXAMPLE
            'Instance1', 'Server1/Instance1', 'Server2' | Invoke-Sqlcmd2 -query "Sp_databases" -as psobject -AppendServerInstance
            This example lists databases for each instance.  It includes a column for the ServerInstance in question.
                DATABASE_NAME          DATABASE_SIZE REMARKS        ServerInstance
                -------------          ------------- -------        --------------
                REDACTED                       88320                Instance1
                master                         17920                Instance1
                ...
                msdb                          618112                Server1/Instance1
                tempdb                        563200                Server1/Instance1
                ...
                OperationsManager           20480000                Server2
        .EXAMPLE
            #Construct a query using SQL parameters
                $Query = "SELECT ServerName, VCServerClass, VCServerContact FROM tblServerInfo WHERE VCServerContact LIKE @VCServerContact AND VCServerClass LIKE @VCServerClass"
            #Run the query, specifying values for SQL parameters
                Invoke-Sqlcmd2 -ServerInstance SomeServer\NamedInstance -Database ServerDB -query $query -SqlParameters @{ VCServerContact="%cookiemonster%"; VCServerClass="Prod" }
                ServerName    VCServerClass VCServerContact
                ----------    ------------- ---------------
                SomeServer1   Prod          cookiemonster, blah
                SomeServer2   Prod          cookiemonster
                SomeServer3   Prod          blah, cookiemonster
        .EXAMPLE
            Invoke-Sqlcmd2 -SQLConnection $Conn -Query "SELECT login_time AS 'StartTime' FROM sysprocesses WHERE spid = 1"
            Uses an existing SQLConnection and runs a basic T-SQL query against it
            StartTime
            -----------
            2010-08-12 21:21:03.593
        .EXAMPLE
            Invoke-SqlCmd2 -SQLConnection $Conn -Query "SELECT ServerName FROM tblServerInfo WHERE ServerName LIKE @ServerName" -SqlParameters @{"ServerName = "c-is-hyperv-1"}
            Executes a parameterized query against the existing SQLConnection, with a collection of one parameter to be passed to the query when executed.
        .EXAMPLE
            Invoke-Sqlcmd2 -ServerInstance "MyComputer\MyInstance" -Query "PRINT 1;SELECT login_time AS 'StartTime' FROM sysprocesses WHERE spid = 1" -Verbose
            Sends "messages" to the Verbose stream, the output stream will hold the results
        .EXAMPLE
            Invoke-Sqlcmd2 -ServerInstance "MyComputer\MyInstance" -Query "PRINT 1;SELECT login_time AS 'StartTime' FROM sysprocesses WHERE spid = 1" -MessagesToOutput
            Sends "messages" to the output stream (irregardless of -Verbose). If you need to "separate" the results, inspecting the type gets really handy:
                    $results = Invoke-Sqlcmd2 -ServerInstance ... -MessagesToOutput
                    $tableResults = $results | Where-Object { $_.GetType().Name -eq 'DataRow' }
                    $messageResults = $results | Where-Object { $_.GetType().Name -ne 'DataRow' }
        .NOTES
            Changelog moved to CHANGELOG.md:
            https://github.com/sqlcollaborative/Invoke-SqlCmd2/blob/master/CHANGELOG.md
        .LINK
            https://github.com/sqlcollaborative/Invoke-SqlCmd2
        .LINK
            https://github.com/RamblingCookieMonster/PowerShell
        .FUNCTIONALITY
            SQL
    #>

    [CmdletBinding(DefaultParameterSetName = 'Ins-Que')]
    [OutputType([System.Management.Automation.PSCustomObject], [System.Data.DataRow], [System.Data.DataTable], [System.Data.DataTableCollection], [System.Data.DataSet])]
    param (
        [Parameter(ParameterSetName = 'Ins-Que',
            Position = 0,
            Mandatory = $true,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true,
            ValueFromRemainingArguments = $false,
            HelpMessage = 'SQL Server Instance required...')]
        [Parameter(ParameterSetName = 'Ins-Fil',
            Position = 0,
            Mandatory = $true,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true,
            ValueFromRemainingArguments = $false,
            HelpMessage = 'SQL Server Instance required...')]
        [Alias('Instance', 'Instances', 'ComputerName', 'Server', 'Servers', 'SqlInstance')]
        [ValidateNotNullOrEmpty()]
        [string[]]$ServerInstance,
        [Parameter(Position = 1,
            Mandatory = $false,
            ValueFromPipelineByPropertyName = $true,
            ValueFromRemainingArguments = $false)]
        [string]$Database,
        [Parameter(ParameterSetName = 'Ins-Que',
            Position = 2,
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true,
            ValueFromRemainingArguments = $false)]
        [Parameter(ParameterSetName = 'Con-Que',
            Position = 2,
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true,
            ValueFromRemainingArguments = $false)]
        [string]$Query,
        [Parameter(ParameterSetName = 'Ins-Fil',
            Position = 2,
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true,
            ValueFromRemainingArguments = $false)]
        [Parameter(ParameterSetName = 'Con-Fil',
            Position = 2,
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true,
            ValueFromRemainingArguments = $false)]
        [ValidateScript( { Test-Path -LiteralPath $_ })]
        [string]$InputFile,
        [Parameter(ParameterSetName = 'Ins-Que',
            Position = 3,
            Mandatory = $false,
            ValueFromPipelineByPropertyName = $true,
            ValueFromRemainingArguments = $false)]
        [Parameter(ParameterSetName = 'Ins-Fil',
            Position = 3,
            Mandatory = $false,
            ValueFromPipelineByPropertyName = $true,
            ValueFromRemainingArguments = $false)]
        [Alias('SqlCredential')]
        [System.Management.Automation.PSCredential]$Credential,
        [Parameter(ParameterSetName = 'Ins-Que',
            Position = 4,
            Mandatory = $false,
            ValueFromRemainingArguments = $false)]
        [Parameter(ParameterSetName = 'Ins-Fil',
            Position = 4,
            Mandatory = $false,
            ValueFromRemainingArguments = $false)]
        [switch]$Encrypt,
        [Parameter(Position = 5,
            Mandatory = $false,
            ValueFromPipelineByPropertyName = $true,
            ValueFromRemainingArguments = $false)]
        [Int32]$QueryTimeout = 600,
        [Parameter(ParameterSetName = 'Ins-Fil',
            Position = 6,
            Mandatory = $false,
            ValueFromPipelineByPropertyName = $true,
            ValueFromRemainingArguments = $false)]
        [Parameter(ParameterSetName = 'Ins-Que',
            Position = 6,
            Mandatory = $false,
            ValueFromPipelineByPropertyName = $true,
            ValueFromRemainingArguments = $false)]
        [Int32]$ConnectionTimeout = 15,
        [Parameter(Position = 7,
            Mandatory = $false,
            ValueFromPipelineByPropertyName = $true,
            ValueFromRemainingArguments = $false)]
        [ValidateSet("DataSet", "DataTable", "DataRow", "PSObject", "SingleValue")]
        [string]$As = "DataRow",
        [Parameter(Position = 8,
            Mandatory = $false,
            ValueFromPipelineByPropertyName = $true,
            ValueFromRemainingArguments = $false)]
        [System.Collections.IDictionary]$SqlParameters,
        [Parameter(Position = 9,
            Mandatory = $false)]
        [switch]$AppendServerInstance,
        [Parameter(Position = 10,
            Mandatory = $false)]
        [switch]$ParseGO,
        [Parameter(ParameterSetName = 'Con-Que',
            Position = 11,
            Mandatory = $false,
            ValueFromPipeline = $false,
            ValueFromPipelineByPropertyName = $false,
            ValueFromRemainingArguments = $false)]
        [Parameter(ParameterSetName = 'Con-Fil',
            Position = 11,
            Mandatory = $false,
            ValueFromPipeline = $false,
            ValueFromPipelineByPropertyName = $false,
            ValueFromRemainingArguments = $false)]
        [Alias('Connection', 'Conn')]
        [ValidateNotNullOrEmpty()]
        [System.Data.SqlClient.SQLConnection]$SQLConnection,
        [Parameter(Position = 12,
            Mandatory = $false)]
        [Alias( 'Application', 'AppName' )]
        [String]$ApplicationName,
        [Parameter(Position = 13,
            Mandatory = $false)]
        [switch]$MessagesToOutput
    )

    begin {
        function Resolve-SqlError {
            param($Err)
            if ($Err) {
                if ($Err.Exception.GetType().Name -eq 'SqlException') {
                    # For SQL exception
                    #$Err = $_
                    Write-Debug -Message "Capture SQL Error"
                    if ($PSBoundParameters.Verbose) {
                        Write-Verbose -Message "SQL Error:  $Err"
                    } #Shiyang, add the verbose output of exception
                    switch ($ErrorActionPreference.ToString()) {
                        { 'SilentlyContinue', 'Ignore' -contains $_ } {   }
                        'Stop' { throw $Err }
                        'Continue' { throw $Err }
                        Default { Throw $Err }
                    }
                }
                else {
                    # For other exception
                    Write-Debug -Message "Capture Other Error"
                    if ($PSBoundParameters.Verbose) {
                        Write-Verbose -Message "Other Error:  $Err"
                    }
                    switch ($ErrorActionPreference.ToString()) {
                        { 'SilentlyContinue', 'Ignore' -contains $_ } { }
                        'Stop' { throw $Err }
                        'Continue' { throw $Err }
                        Default { throw $Err }
                    }
                }
            }

        }
        if ($InputFile) {
            $filePath = $(Resolve-Path -LiteralPath $InputFile).ProviderPath
            $Query = [System.IO.File]::ReadAllText("$filePath")
        }

        Write-Debug -Message "Running Invoke-Sqlcmd2 with ParameterSet '$($PSCmdlet.ParameterSetName)'.  Performing query '$Query'."

        if ($As -eq "PSObject") {
            #This code scrubs DBNulls.  Props to Dave Wyatt
            $cSharp = @'
                using System;
                using System.Data;
                using System.Management.Automation;
                public class DBNullScrubber
                {
                    public static PSObject DataRowToPSObject(DataRow row)
                    {
                        PSObject psObject = new PSObject();
                        if (row != null && (row.RowState & DataRowState.Detached) != DataRowState.Detached)
                        {
                            foreach (DataColumn column in row.Table.Columns)
                            {
                                Object value = null;
                                if (!row.IsNull(column))
                                {
                                    value = row[column];
                                }
                                psObject.Properties.Add(new PSNoteProperty(column.ColumnName, value));
                            }
                        }
                        return psObject;
                    }
                }
'@

            try {
                if ($PSEdition -ne 'Core'){
                    Add-Type -TypeDefinition $cSharp -ReferencedAssemblies 'System.Data', 'System.Xml' -ErrorAction stop
                } else {
                    Add-Type $cSharp -ErrorAction stop
                }

                
            }
            catch {
                if (-not $_.ToString() -like "*The type name 'DBNullScrubber' already exists*") {
                    Write-Warning "Could not load DBNullScrubber.  Defaulting to DataRow output: $_."
                    $As = "Datarow"
                }
            }
        }

        #Handle existing connections
        if ($PSBoundParameters.ContainsKey('SQLConnection')) {
            if ($SQLConnection.State -notlike "Open") {
                try {
                    Write-Debug -Message "Opening connection from '$($SQLConnection.State)' state."
                    $SQLConnection.Open()
                }
                catch {
                    throw $_
                }
            }

            if ($Database -and $SQLConnection.Database -notlike $Database) {
                try {
                    Write-Debug -Message "Changing SQLConnection database from '$($SQLConnection.Database)' to $Database."
                    $SQLConnection.ChangeDatabase($Database)
                }
                catch {
                    throw "Could not change Connection database '$($SQLConnection.Database)' to $Database`: $_"
                }
            }

            if ($SQLConnection.state -like "Open") {
                $ServerInstance = @($SQLConnection.DataSource)
            }
            else {
                throw "SQLConnection is not open"
            }
        }
        $GoSplitterRegex = [regex]'(?smi)^[\s]*GO[\s]*$'

    }
    process {
        foreach ($SQLInstance in $ServerInstance) {
            Write-Debug -Message "Querying ServerInstance '$SQLInstance'"

            if ($PSBoundParameters.Keys -contains "SQLConnection") {
                $Conn = $SQLConnection
            }
            else {
                $CSBuilder = New-Object -TypeName System.Data.SqlClient.SqlConnectionStringBuilder
                $CSBuilder["Server"] = $SQLInstance
                $CSBuilder["Database"] = $Database
                $CSBuilder["Connection Timeout"] = $ConnectionTimeout

                if ($Encrypt) {
                    $CSBuilder["Encrypt"] = $true
                }

                if ($Credential) {
                    $CSBuilder["Trusted_Connection"] = $false
                    $CSBuilder["User ID"] = $Credential.UserName
                    $CSBuilder["Password"] = $Credential.GetNetworkCredential().Password
                }
                else {
                    $CSBuilder["Integrated Security"] = $true
                }
                if ($ApplicationName) {
                    $CSBuilder["Application Name"] = $ApplicationName
                }
                else {
                    $ScriptName = (Get-PSCallStack)[-1].Command.ToString()
                    if ($ScriptName -ne "<ScriptBlock>") {
                        $CSBuilder["Application Name"] = $ScriptName
                    }
                }
                $conn = New-Object -TypeName System.Data.SqlClient.SQLConnection

                $ConnectionString = $CSBuilder.ToString()
                $conn.ConnectionString = $ConnectionString
                Write-Debug "ConnectionString $ConnectionString"

                try {
                    $conn.Open()
                }
                catch {
                    Write-Error $_
                    continue
                }
            }


            if ($ParseGO) {
                Write-Debug -Message "Stripping GOs from source"
                $Pieces = $GoSplitterRegex.Split($Query)
            }
            else {
                $Pieces = , $Query
            }
            # Only execute non-empty statements
            $Pieces = $Pieces | Where-Object { $_.Trim().Length -gt 0 }
            foreach ($piece in $Pieces) {
                $cmd = New-Object system.Data.SqlClient.SqlCommand($piece, $conn)
                $cmd.CommandTimeout = $QueryTimeout

                if ($null -ne $SqlParameters) {
                    $SqlParameters.GetEnumerator() |
                        ForEach-Object {
                        if ($null -ne $_.Value) {
                            $cmd.Parameters.AddWithValue($_.Key, $_.Value)
                        }
                        else {
                            $cmd.Parameters.AddWithValue($_.Key, [DBNull]::Value)
                        }
                    } > $null
                }

                $ds = New-Object system.Data.DataSet
                $da = New-Object system.Data.SqlClient.SqlDataAdapter($cmd)

                if ($MessagesToOutput) {
                    $pool = [RunspaceFactory]::CreateRunspacePool(1, [int]$env:NUMBER_OF_PROCESSORS + 1)
                    $pool.ApartmentState = "MTA"
                    $pool.Open()
                    $runspaces = @()
                    $scriptblock = {
                        Param ($da, $ds, $conn, $queue )
                        $conn.FireInfoMessageEventOnUserErrors = $false
                        $handler = [System.Data.SqlClient.SqlInfoMessageEventHandler] { $queue.Enqueue($_) }
                        $conn.add_InfoMessage($handler)
                        $Err = $null
                        try {
                            [void]$da.fill($ds)
                        }
                        catch {
                            $Err = $_
                        }
                        finally {
                            $conn.remove_InfoMessage($handler)
                        }
                        return $Err
                    }
                    $queue = New-Object System.Collections.Concurrent.ConcurrentQueue[string]
                    $runspace = [PowerShell]::Create()
                    $null = $runspace.AddScript($scriptblock)
                    $null = $runspace.AddArgument($da)
                    $null = $runspace.AddArgument($ds)
                    $null = $runspace.AddArgument($Conn)
                    $null = $runspace.AddArgument($queue)
                    $runspace.RunspacePool = $pool
                    $runspaces += [PSCustomObject]@{ Pipe = $runspace; Status = $runspace.BeginInvoke() }
                    # While streaming ...
                    while ($runspaces.Status.IsCompleted -notcontains $true) {
                        $item = $null
                        if ($queue.TryDequeue([ref]$item)) {
                            "$item"
                        }
                    }
                    # Drain the stream as the runspace is closed, just to be safe
                    if ($queue.IsEmpty -ne $true) {
                        $item = $null
                        while ($queue.TryDequeue([ref]$item)) {
                            "$item"
                        }
                    }
                    foreach ($runspace in $runspaces) {
                        $results = $runspace.Pipe.EndInvoke($runspace.Status)
                        $runspace.Pipe.Dispose()
                        if ($null -ne $results) {
                            Resolve-SqlError $results[0]
                        }
                    }
                    $pool.Close()
                    $pool.Dispose()
                }
                else {
                    #Following EventHandler is used for PRINT and RAISERROR T-SQL statements. Executed when -Verbose parameter specified by caller and no -MessageToOutput
                    if ($PSBoundParameters.Verbose) {
                        $conn.FireInfoMessageEventOnUserErrors = $false
                        $handler = [System.Data.SqlClient.SqlInfoMessageEventHandler] { Write-Verbose "$($_)" }
                        $conn.add_InfoMessage($handler)
                    }
                    try {
                        [void]$da.fill($ds)
                    }
                    catch {
                        $Err = $_
                    }
                    finally {
                        if ($PSBoundParameters.Verbose) {
                            $conn.remove_InfoMessage($handler)
                        }
                    }
                    Resolve-SqlError $Err
                }
                #Close the connection
                if (-not $PSBoundParameters.ContainsKey('SQLConnection')) {
                    $Conn.Close()
                }
                if ($AppendServerInstance) {
                    #Basics from Chad Miller
                    $Column = New-Object Data.DataColumn
                    $Column.ColumnName = "ServerInstance"

                    if ($ds.Tables.Count -ne 0) {
                        $ds.Tables[0].Columns.Add($Column)
                        Foreach ($row in $ds.Tables[0]) {
                            $row.ServerInstance = $SQLInstance
                        }
                    }
                }

                switch ($As) {
                    'DataSet' {
                        $ds
                    }
                    'DataTable' {
                        $ds.Tables
                    }
                    'DataRow' {
                        if ($ds.Tables.Count -ne 0) {
                            $ds.Tables[0]
                        }
                    }
                    'PSObject' {
                        if ($ds.Tables.Count -ne 0) {
                            #Scrub DBNulls - Provides convenient results you can use comparisons with
                            #Introduces overhead (e.g. ~2000 rows w/ ~80 columns went from .15 Seconds to .65 Seconds - depending on your data could be much more!)
                            foreach ($row in $ds.Tables[0].Rows) {
                                [DBNullScrubber]::DataRowToPSObject($row)
                            }
                        }
                    }
                    'SingleValue' {
                        if ($ds.Tables.Count -ne 0) {
                            $ds.Tables[0] | Select-Object -ExpandProperty $ds.Tables[0].Columns[0].ColumnName
                        }
                    }
                }
            } #foreach ($piece in $Pieces)
        }
    }
} #Invoke-Sqlcmd2

function Get-ADSIObject {  
    <#
    .SYNOPSIS
	    Get AD object (user, group, etc.) via ADSI.
    .DESCRIPTION
	    Get AD object (user, group, etc.) via ADSI.
        Invoke a specify an LDAP Query, or search based on samaccountname and/or objectcategory
    .FUNCTIONALITY
        Active Directory
    .PARAMETER samAccountName
        Specific samaccountname to filter on
    .PARAMETER ObjectCategory
        Specific objectCategory to filter on
    
    .PARAMETER Query
        LDAP filter to invoke
    .PARAMETER Path
        LDAP Path.  e.g. contoso.com, DomainController1
        LDAP:// is prepended when omitted
    .PARAMETER Property
        Specific properties to query for
 
    .PARAMETER Limit
        If specified, limit results to this size
    .PARAMETER Credential
        Credential to use for query
        If specified, the Path parameter must be specified as well.
    .PARAMETER As
        SearchResult        = results directly from DirectorySearcher
        DirectoryEntry      = Invoke GetDirectoryEntry against each DirectorySearcher object returned
        PSObject (Default)  = Create a PSObject with expected properties and types
    .EXAMPLE
        Get-ADSIObject jdoe
        # Find an AD object with the samaccountname jdoe
    .EXAMPLE
        Get-ADSIObject -Query "(&(objectCategory=Group)(samaccountname=domain admins))"
        # Find an AD object meeting the specified criteria
    .EXAMPLE
        Get-ADSIObject -Query "(objectCategory=Group)" -Path contoso.com
        # List all groups at the root of contoso.com
    
    .EXAMPLE
        Echo jdoe, cmonster | Get-ADSIObject -property mail -ObjectCategory User | Select -expandproperty mail
        # Find an AD object for a few users, extract the mail property only
    .EXAMPLE
        $DirectoryEntry = Get-ADSIObject TESTUSER -as DirectoryEntry
        $DirectoryEntry.put(‘Title’,’Test’) 
        $DirectoryEntry.setinfo()
        #Get the AD object for TESTUSER in a usable form (DirectoryEntry), set the title attribute to Test, and make the change.
    .LINK
        https://gallery.technet.microsoft.com/scriptcenter/Get-ADSIObject-Portable-ae7f9184
    #>	
    [cmdletbinding(DefaultParameterSetName='SAM')]
    Param(
        [Parameter( Position=0,
                    Mandatory = $true,
                    ValueFromPipeline=$true,
                    ValueFromPipelineByPropertyName=$true,
                    ParameterSetName='SAM')]
        [string[]]$samAccountName = "*",

        [Parameter( Position=1,
                    ParameterSetName='SAM')]
        [string[]]$ObjectCategory = "*",

        [Parameter( ParameterSetName='Query',
                    Mandatory = $true )]
        [string]$Query = $null,

        [string]$Path = $Null,

        [string[]]$Property = $Null,

        [int]$Limit,

        [System.Management.Automation.PSCredential]$Credential,

        [validateset("PSObject","DirectoryEntry","SearchResult")]
        [string]$As = "PSObject"
    )

    Begin 
    {
        #Define parameters for creating the object
        $Params = @{
            TypeName = "System.DirectoryServices.DirectoryEntry"
            ErrorAction = "Stop"
        }

        #If we have an LDAP path, add it in.
            if($Path){

                if($Path -notlike "^LDAP")
                {
                    $Path = "LDAP://$Path"
                }
            
                $Params.ArgumentList = @($Path)

                #if we have a credential, add it in
                if($Credential)
                {
                    $Params.ArgumentList += $Credential.UserName
                    $Params.ArgumentList += $Credential.GetNetworkCredential().Password
                }
            }
            elseif($Credential)
            {
                Throw "Using the Credential parameter requires a valid Path parameter"
            }

        #Create the domain entry for search root
            Try
            {
                Write-Verbose "Bound parameters:`n$($PSBoundParameters | Format-List | Out-String )`nCreating DirectoryEntry with parameters:`n$($Params | Out-String)"
                $DomainEntry = New-Object @Params
            }
            Catch
            {
                Throw "Could not establish DirectoryEntry: $_"
            }
            $DomainName = $DomainEntry.name

        #Set up the searcher
            $Searcher = New-Object -TypeName System.DirectoryServices.DirectorySearcher
            $Searcher.PageSize = 1000
            $Searcher.SearchRoot = $DomainEntry
            if($Limit)
            {
                $Searcher.SizeLimit = $limit
            }
            if($Property)
            {
                foreach($Prop in $Property)
                {
                    $Searcher.PropertiesToLoad.Add($Prop) | Out-Null
                }
            }




		#Define a function to get ADSI results from a specific query
        Function Get-ADSIResult
        {
            [cmdletbinding()]
            param(
                [string[]]$Property = $Null,
                [string]$Query,
                [string]$As,
                $Searcher
            )
            
            #Invoke the query
                $Results = $null
                $Searcher.Filter = $Query
                $Results = $Searcher.FindAll()
            
            #If SearchResult, just spit out the results.
                if($As -eq "SearchResult")
                {
                    $Results
                }
            #If DirectoryEntry, invoke GetDirectoryEntry
                elseif($As -eq "DirectoryEntry")
                {
                    $Results | ForEach-Object { $_.GetDirectoryEntry() }
                }
            #Otherwise, get properties from the object
                else
                {
                    $Results | ForEach-Object {
                
                        #Get the keys.  They aren't an array, so split them up, remove empty, and trim just in case I screwed something up...
                            $object = $_
                            #cast to array of strings or else PS2 breaks when we select down the line
                            [string[]]$properties = ($object.properties.PropertyNames) -split "`r|`n" | Where-Object { $_ } | ForEach-Object { $_.Trim() }
            
                        #Filter properties if desired
                            if($Property)
                            {
                                $properties = $properties | Where-Object {$Property -Contains $_}
                            }
            
                        #Build up an object to output.  Loop through each property, extract from ResultPropertyValueCollection
                            #Create the object, PS2 compatibility.  can't just pipe to select, props need to exist
                                $hash = @{}
                                foreach($prop in $properties)
                                {
                                    $hash.$prop = $null
                                }
                                $Temp = New-Object -TypeName PSObject -Property $hash | Select -Property $properties
                        
                            foreach($Prop in $properties)
                            {
                                Try
                                {
                                    $Temp.$Prop = foreach($item in $object.properties.$prop)
                                    {
                                        $item
                                    }
                                }
                                Catch
                                {
                                    Write-Warning "Could not get property '$Prop': $_"
                                }   
                            }
                            $Temp
                    }
                }
        }
    }
    Process
    {
        #Set up the query as defined, or look for a samaccountname.  Probably a cleaner way to do this...
            if($PsCmdlet.ParameterSetName -eq 'Query'){
                Write-Verbose "Working on Query '$Query'"
                Get-ADSIResult -Searcher $Searcher -Property $Property -Query $Query -As $As
            }
            else
            {
                foreach($AccountName in $samAccountName)
                {
                    #Build up the LDAP query...
                        $QueryArray = @( "(samAccountName=$AccountName)" )
                        if($ObjectCategory)
                        {
                            [string]$TempString = ( $ObjectCategory | ForEach-Object {"(objectCategory=$_)"} ) -join ""
                            $QueryArray += "(|$TempString)"
                        }
                        $Query = "(&$($QueryArray -join ''))"
                    Write-Verbose "Working on built Query '$Query'"
                    Get-ADSIResult -Searcher $Searcher -Property $Property -Query $Query -As $As
                }
            }
    }
    End
    {
        $Searcher = $null
        $DomainEntry = $null
    }
}


# Load a xml file
function LoadXml ($filename)
{
    $XamlLoader=(New-Object System.Xml.XmlDocument)
    $XamlLoader.Load($filename)
    return $XamlLoader
}

####################################
##### Import Session Functions
####################################
# Import functions from the current session into the RunspacePool sessionstate

# Create runspace session state
$InitialSessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()

# Import all session functions into the runspace session state from the current one
Get-ChildItem Function:\ | Where-Object {$_.name -notlike "*:*"} |  select name -ExpandProperty name |
ForEach-Object {       

    # Get the function code
    $Definition = Get-Content "function:\$_" -ErrorAction Stop

    # Create a sessionstate function with the same name and code
    $SessionStateFunction = New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList "$_", $Definition

    # Add the function to the session state
    $InitialSessionState.Commands.Add($SessionStateFunction)
}

####################################
##### Runspace 1 Init
####################################

$Global:syncHash = [hashtable]::Synchronized(@{})
$newRunspace =[runspacefactory]::CreateRunspace($InitialSessionState)
$newRunspace.ApartmentState = "STA"
$newRunspace.ThreadOptions = "ReuseThread"         
$newRunspace.Open()
$newRunspace.SessionStateProxy.SetVariable("syncHash",$Global:syncHash) 

$Global:synchash.location = $global:getloaction
####################################
##### Runspace 1 Script
####################################
$psCmd = [PowerShell]::Create().AddScript({
    #Output errors to hashtable
    $Global:syncHash.Error = $Error

    # Load MainWindow
    $XamlMainWindow=LoadXml("$($Global:synchash.location)\mahapps.xaml")
    $Reader=(New-Object System.Xml.XmlNodeReader $XamlMainWindow)
    $Global:syncHash.Form=[Windows.Markup.XamlReader]::Load($Reader)

    $CompName = $env:COMPUTERNAME
    $object = New-Object -comObject Shell.Application  
    ####################################
    ##### Buttons initialization
    ####################################
    $Global:syncHash.choose_server = $Global:syncHash.Form.Findname("choose_server")
    $Global:syncHash.statsbox = $Global:syncHash.Form.Findname("statsbox")

    ####################################
    ##### Dropdown group initialization
    ####################################
    $servergroups = Get-ADSIObject -Query "(objectCategory=Group)" -Path "OU=Server Patching,OU=Machine Groups,OU=Groups,DC=svhanational,DC=org,DC=au" | Where-Object {($_.name -like "su_Servers_MaintWindow_*" -or $_.name -like "su_Servers_DeployType_RebootSuppressed_ASAP*")}

    foreach($server in $servergroups){
        $name = $server.name
        $null = $Global:syncHash.choose_server.Items.Add($name)
    }
    ####################################
    ##### Form Actions
    ####################################
    $Global:syncHash.Form.Add_Closing({
        ([System.Diagnostics.Process]::GetCurrentProcess() | Get-Process) | Stop-Process
    })
    $Global:syncHash.choose_server.Add_DropDownClosed({
        $Global:syncHash.SQLCHECK = 0
        $Global:syncHash.statsbox.items.clear()

        #Write-Host $choose_server.text
        $group = $Global:syncHash.choose_server.text


######################################################################################################################################################################################################################
$names = @()
$members = @()
#        $members = Get-ADGroupMember $group
$Groupe = [ADSI]"LDAP://CN=$group,OU=Server Patching,OU=Machine Groups,OU=Groups,DC=svhanational,DC=org,DC=au"
$Groupe.Member | ForEach-Object {
    $Searcher = [adsisearcher]"(distinguishedname=$_)"
    $names += $searcher.FindOne().Properties.cn 

}
$names = $names  | Sort-Object
foreach($name in $names){
    Write-Host "foreach $name"
    If($name -Match "S-1-5"){
    $convertedname = Convert-FspToUsername -UserSID $name
        $name = $convertedname.sAMAccountName
        
    If($convertedname.sAMAccountName -match "\\"){
           $splitname = $convertedname.sAMAccountName.split("\")
            $name = $splitname[1]
        }
    }
    If($name -match '$'){
        $name = $name.trimend('$')
    }
$members += $name
}
$members

######################################################################################################################################################################################################################

        foreach($member in $members){
            #Write-Host $member.name
            $hn = $member

            $Query = "
            SELECT Hostname, UpdateStatus, LastUpdateTime FROM dbo.Clients WHERE Hostname='"+$hn+"'
                "
            
            $hostname = Invoke-Sqlcmd2 -ServerInstance "cmtools-pr9-01\clienthealth" -Database 'ClientHealth' -Query $query -as PSObject
                  
            $Date = Get-Date
            if(!$hostname) {
                $hostname = @{
                    hostname = "$hn"
                    UpdateStatus = "Not Communicating with SQL"
                }        
            }
            elseif(!$hostname.LastUpdateTime){
                $hostname = @{
                    hostname = "$hn"
                    UpdateStatus = "No Update Data"
                    LastUpdateTime = $null
                }
            }
            elseif($hostname.LastUpdateTime -lt $Date.AddDays(-2)){
                $usdate = $hostname.LastUpdateTime
                $dateParts = $usdate -split "/"
                $deDate = "$($dateparts[1])/$($dateParts[0])/$($dateParts[2])"
                $deDate
                $hostname = @{
                    hostname = "$hn"
                    UpdateStatus = "Last updated +2 days ago"
                    LastUpdateTime = $deDate
                }        
            }
            elseif($hostname.LastUpdateTime -ne $null){
                $usdate = $hostname.LastUpdateTime
                $dateParts = $usdate -split "/"
                $deDate = "$($dateparts[1])/$($dateParts[0])/$($dateParts[2])"
                $deDate
                $hostname = @{
                    hostname = "$hn"
                    UpdateStatus = $hostname.UpdateStatus
                    LastUpdateTime = $deDate
                }
            }
            $Global:syncHash.statsbox.items.Add([pscustomobject]$hostname)
        }
        $Global:syncHash.SQLCHECK = 1
    })
    $Global:syncHash.Form.ShowDialog() | Out-Null
})
####################################
##### Runspace 2 Init
####################################

$psCmd.Runspace = $newRunspace
$newRunspace2 =[runspacefactory]::CreateRunspace($InitialSessionState)
$newRunspace2.ApartmentState = "STA"
$newRunspace2.ThreadOptions = "ReuseThread"         
$newRunspace2.Open()
$newRunspace2.SessionStateProxy.SetVariable("syncHash",$Global:syncHash) 

####################################
##### Runspace 2 Script
####################################
$psCmd2 = [PowerShell]::Create().AddScript({ 
while($true){
    if($Global:syncHash.SQLCHECK -eq 1){
        $members = $Global:syncHash.statsbox.items
        foreach($member in $members){

            #Write-Host $member.name
            $hn = $member.hostname
        
            $Query = "
            SELECT Hostname, UpdateStatus, LastUpdateTime FROM dbo.Clients WHERE Hostname='"+$hn+"'
                "
            
            $hostname = Invoke-Sqlcmd2 -ServerInstance "cmtools-pr9-01\clienthealth" -Database 'ClientHealth' -Query $query -as PSObject
        
            $Date = Get-Date 

            if(!$hostname) {
                $hostname = @{
                    hostname = "$hn"
                    UpdateStatus = "Not Communicating with SQL"
                }        
            }
            elseif(!$hostname.LastUpdateTime){
                $hostname = @{
                    hostname = "$hn"
                    UpdateStatus = "No Update Data"
                    LastUpdateTime = $null
                }
            }
            elseif($hostname.LastUpdateTime -lt $Date.AddDays(-2)){
                $usdate = $hostname.LastUpdateTime
                $dateParts = $usdate -split "/"
                $deDate = "$($dateparts[1])/$($dateParts[0])/$($dateParts[2])"
                $deDate
                $hostname = @{
                    hostname = "$hn"
                    UpdateStatus = "Last updated +2 days ago"
                    LastUpdateTime = $deDate
                }        
            }
            elseif($hostname.LastUpdateTime -ne $null){
                $usdate = $hostname.LastUpdateTime
                $dateParts = $usdate -split "/"
                $deDate = "$($dateparts[1])/$($dateParts[0])/$($dateParts[2])"
                $deDate
                $hostname = @{
                    hostname = "$hn"
                    UpdateStatus = $hostname.UpdateStatus
                    LastUpdateTime = $deDate
                }
            }
            $member.updatestatus = $hostname.updatestatus
            $member.lastupdatetime = $hostname.LastUpdateTime            
        }   
        $Global:syncHash.form.Dispatcher.Invoke([action]{$Global:syncHash.statsbox.items.Refresh()},"Normal")
    }
    start-sleep -seconds 10
}

})
$psCmd2.Runspace = $newRunspace2
####################################
##### Runspace Invokes
####################################
$data2 = $psCmd2.BeginInvoke()
$data = $psCmd.BeginInvoke()
