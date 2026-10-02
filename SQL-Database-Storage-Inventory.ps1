<#
================================================================================
 SQL Server Database Storage Inventory
================================================================================

 Author   : Russell McKee
 Copyright: Copyright (C) 2026 Russell McKee
 SPDX-License-Identifier: GPL-3.0-only
 LinkedIn : https://www.linkedin.com/in/russellwbmckee/
 Version  : 1.3
 Updated  : 2 October 2026

.SYNOPSIS
    Reports database counts and actual used database storage for SQL Server
    instances listed in the SQL Server Discovery Toolkit CSV outputs.

.DESCRIPTION
    Imports one or more CSV files produced by SQL-Discovery-DomainController.ps1
    or SQL-Discovery-MemberServer.ps1, deduplicates the discovered instances,
    and connects to each instance from a Windows member server.

    By default, connections use Windows integrated authentication with the
    identity running PowerShell. Set -AuthMode Prompt or supply -Credential
    when SQL Server authentication is required. An explicitly supplied
    credential takes precedence over -AuthMode.

    Actual data usage is calculated from FILEPROPERTY(name, 'SpaceUsed') for
    every accessible online database. Actual transaction-log usage is obtained
    from DBCC SQLPERF(LOGSPACE). Results include system databases unless
    -ExcludeSystemDatabases is specified.

.PARAMETER InputCsv
    One or more discovery CSV paths. By default, the script looks for both
    standard discovery output files in the current directory.

.PARAMETER OutputCsv
    Destination CSV path.

.PARAMETER DiagnosticLog
    Append-only diagnostic text file path. Console errors include a diagnostic
    ID that identifies the corresponding detailed entry. Existing content is
    preserved. The file is created or appended only when an error or warning is
    captured. Error details are not written to the inventory CSV.

.PARAMETER AuthMode
    CurrentUser : use Windows integrated authentication with the identity
                  running PowerShell (default).
    Prompt      : securely prompt for a SQL Server login and password.

.PARAMETER Credential
    Optional SQL Server authentication credential supplied directly. It takes
    precedence over -AuthMode. -SqlCredential remains available as an alias.
    This does not impersonate an alternate Windows account.

.PARAMETER ExcludeSystemDatabases
    Excludes master, model, msdb and tempdb from the count and storage totals.

.PARAMETER ConnectionTimeoutSeconds
    SQL connection timeout in seconds.

.PARAMETER CommandTimeoutSeconds
    SQL command timeout in seconds for each inventory query.

.PARAMETER Encrypt
    Requests an encrypted SQL connection.

.PARAMETER TrustServerCertificate
    Trusts the server certificate when encryption is enabled. Use only when
    permitted by organizational policy.

.EXAMPLE
    .\SQL-Database-Storage-Inventory.ps1

.EXAMPLE
    .\SQL-Database-Storage-Inventory.ps1 `
        -InputCsv '.\SQL-Discovery-DomainController.csv',
                  '.\SQL-Discovery-MemberServer.csv' `
        -OutputCsv 'C:\Reports\SQL-Database-Storage.csv'

.EXAMPLE
    $credential = Get-Credential
    .\SQL-Database-Storage-Inventory.ps1 -Credential $credential

.EXAMPLE
    .\SQL-Database-Storage-Inventory.ps1 -AuthMode Prompt

.NOTES
    REQUIREMENTS
      - Windows PowerShell 5.1 or PowerShell 7.
      - Network access to every discovered SQL Server endpoint.
      - A login that can connect to each database to be measured.
      - Permission to run DBCC SQLPERF(LOGSPACE).

    Failed instances remain in the output CSV with FAILED status. Detailed
    errors are appended to the diagnostic text file only when errors occur.

 -------------------------------------------------------------------------------
 DISCLAIMER - PLEASE READ BEFORE RUNNING
 -------------------------------------------------------------------------------
 This script is provided "AS IS", without warranty of any kind, express or
 implied, including but not limited to the warranties of merchantability,
 fitness for a particular purpose and non-infringement.

 In no event shall the author be liable for any claim, damages, loss of data,
 loss of service, or other liability, whether in an action of contract, tort or
 otherwise, arising from, out of, or in connection with the use of this script.

 You run this script entirely at your own risk. It is NOT supported by Microsoft
 and does not constitute official Microsoft guidance.

 ALWAYS TEST IN A NON-PRODUCTION / TEST ENVIRONMENT FIRST and satisfy yourself
 that the behaviour is appropriate before running against production systems.

 DATABASE QUERY NOTICE: only connect to and query SQL Server instances you are
 explicitly authorised to access. Use a dedicated least-privilege login where
 possible. Unauthorised access may breach policy or law and may trigger security
 monitoring and alerting.
 -------------------------------------------------------------------------------
================================================================================
#>

[CmdletBinding()]
param(
    [string[]] $InputCsv = @(
        '.\SQL-Discovery-DomainController.csv',
        '.\SQL-Discovery-MemberServer.csv'
    ),

    [string] $OutputCsv = '.\SQL-Database-Storage-Inventory.csv',

    [string] $DiagnosticLog = '.\SQL-Database-Storage-Inventory-Diagnostics.txt',

    [ValidateSet('CurrentUser', 'Prompt')]
    [string] $AuthMode = 'CurrentUser',

    [Alias('SqlCredential')]
    [System.Management.Automation.PSCredential]
    $Credential = [System.Management.Automation.PSCredential]::Empty,

    [switch] $ExcludeSystemDatabases,

    [ValidateRange(1, 300)]
    [int] $ConnectionTimeoutSeconds = 10,

    [ValidateRange(1, 3600)]
    [int] $CommandTimeoutSeconds = 30,

    [switch] $Encrypt,

    [switch] $TrustServerCertificate
)

$ErrorActionPreference = 'Stop'
$script:StartTime = Get-Date

if ($Credential -eq [System.Management.Automation.PSCredential]::Empty) {
    $Credential = $null
}

if ($AuthMode -eq 'Prompt' -and -not $Credential) {
    $Credential = Get-Credential -Message 'Enter a SQL Server login and password for the storage inventory'
    if (-not $Credential) {
        throw 'Credential prompt was cancelled. No inventory was performed.'
    }
}

function Get-PropertyValue {
    param(
        [psobject] $InputObject,
        [string] $Name
    )

    $property = $InputObject.PSObject.Properties[$Name]
    if ($property) { return [string] $property.Value }
    return ''
}

function Get-SqlDataSource {
    param([psobject] $DiscoveryRow)

    $fqdn = Get-PropertyValue -InputObject $DiscoveryRow -Name 'FQDN'
    $serverName = Get-PropertyValue -InputObject $DiscoveryRow -Name 'ServerName'
    $ipAddress = Get-PropertyValue -InputObject $DiscoveryRow -Name 'IPAddress'
    $instanceName = Get-PropertyValue -InputObject $DiscoveryRow -Name 'InstanceName'
    $detectedBy = Get-PropertyValue -InputObject $DiscoveryRow -Name 'DetectedBy'
    $spns = Get-PropertyValue -InputObject $DiscoveryRow -Name 'ServicePrincipalName'

    $hostName = $fqdn
    if ([string]::IsNullOrWhiteSpace($hostName)) { $hostName = $serverName }
    if ([string]::IsNullOrWhiteSpace($hostName)) { $hostName = $ipAddress }

    if ([string]::IsNullOrWhiteSpace($hostName) -or
        [string]::IsNullOrWhiteSpace($instanceName) -or
        $instanceName -like '(*') {
        return $null
    }

    if ($instanceName -ne 'MSSQLSERVER') {
        return "$hostName\$instanceName"
    }

    if ($detectedBy -match '(?i)\bTCP port\s+(?<Port>\d{1,5})\b') {
        return "tcp:$hostName,$($Matches.Port)"
    }

    foreach ($spn in ($spns -split '\s*;\s*')) {
        if ($spn -match '(?i)^MSSQLSvc/[^:]+:(?<Port>\d{1,5})$') {
            return "tcp:$hostName,$($Matches.Port)"
        }
    }

    return $hostName
}

function Get-InstanceKey {
    param([psobject] $DiscoveryRow)

    $serverName = Get-PropertyValue -InputObject $DiscoveryRow -Name 'ServerName'
    $fqdn = Get-PropertyValue -InputObject $DiscoveryRow -Name 'FQDN'
    $ipAddress = Get-PropertyValue -InputObject $DiscoveryRow -Name 'IPAddress'
    $instanceName = Get-PropertyValue -InputObject $DiscoveryRow -Name 'InstanceName'

    $hostIdentity = $serverName
    if ([string]::IsNullOrWhiteSpace($hostIdentity)) { $hostIdentity = $fqdn }
    if ([string]::IsNullOrWhiteSpace($hostIdentity)) { $hostIdentity = $ipAddress }

    return ('{0}|{1}' -f $hostIdentity.Trim(), $instanceName.Trim())
}

function New-SqlConnectionString {
    param(
        [string] $DataSource,
        [string] $Database
    )

    $builder = [System.Data.SqlClient.SqlConnectionStringBuilder]::new()
    $builder['Data Source'] = $DataSource
    $builder['Initial Catalog'] = $Database
    $builder['Connect Timeout'] = $ConnectionTimeoutSeconds
    $builder['Application Name'] = 'SQL Database Storage Inventory'
    $builder['Encrypt'] = [bool] $Encrypt
    $builder['TrustServerCertificate'] = [bool] $TrustServerCertificate

    if ($Credential) {
        $builder['Integrated Security'] = $false
        $builder['User ID'] = $Credential.UserName
        $builder['Password'] = $Credential.GetNetworkCredential().Password
    }
    else {
        $builder['Integrated Security'] = $true
    }

    return $builder.ConnectionString
}

function Invoke-SqlDataTable {
    param(
        [System.Data.SqlClient.SqlConnection] $Connection,
        [string] $Query
    )

    $command = $Connection.CreateCommand()
    $command.CommandText = $Query
    $command.CommandTimeout = $CommandTimeoutSeconds

    $table = [System.Data.DataTable]::new()
    $reader = $null
    try {
        $reader = $command.ExecuteReader()
        $table.Load($reader)
    }
    finally {
        if ($reader) { $reader.Dispose() }
        $command.Dispose()
    }

    return ,$table
}

function Format-ErrorRecord {
    param([System.Management.Automation.ErrorRecord] $ErrorRecord)

    $lines = [System.Collections.Generic.List[string]]::new()
    $lines.Add("Message: $($ErrorRecord.Exception.Message)")
    $lines.Add("Exception type: $($ErrorRecord.Exception.GetType().FullName)")
    $lines.Add("Fully qualified error ID: $($ErrorRecord.FullyQualifiedErrorId)")
    $lines.Add("Category: $($ErrorRecord.CategoryInfo)")

    if ($ErrorRecord.InvocationInfo -and $ErrorRecord.InvocationInfo.PositionMessage) {
        $lines.Add("Invocation: $($ErrorRecord.InvocationInfo.PositionMessage.Trim())")
    }
    if ($ErrorRecord.ScriptStackTrace) {
        $lines.Add("PowerShell stack trace: $($ErrorRecord.ScriptStackTrace)")
    }

    $exception = $ErrorRecord.Exception
    $depth = 0
    while ($exception) {
        $prefix = if ($depth -eq 0) { 'Exception' } else { "Inner exception $depth" }
        $lines.Add("$prefix type: $($exception.GetType().FullName)")
        $lines.Add("$prefix message: $($exception.Message)")
        if ($exception.StackTrace) {
            $lines.Add("$prefix .NET stack trace: $($exception.StackTrace)")
        }

        if ($exception -is [System.Data.SqlClient.SqlException]) {
            foreach ($sqlError in $exception.Errors) {
                $lines.Add(
                    "SQL error: Number=$($sqlError.Number); State=$($sqlError.State); " +
                    "Class=$($sqlError.Class); Server=$($sqlError.Server); " +
                    "Procedure=$($sqlError.Procedure); Line=$($sqlError.LineNumber); " +
                    "Message=$($sqlError.Message)"
                )
            }
        }

        $exception = $exception.InnerException
        $depth++
    }

    return $lines -join [Environment]::NewLine
}

function Write-Diagnostic {
    param(
        [ValidateSet('INFO', 'WARNING', 'ERROR')]
        [string] $Level,
        [string] $Context,
        [string] $Message,
        [System.Management.Automation.ErrorRecord] $ErrorRecord
    )

    $diagnosticId = [Guid]::NewGuid().ToString('N').Substring(0, 12)
    $lines = [System.Collections.Generic.List[string]]::new()
    $lines.Add('--------------------------------------------------------------------------------')
    $lines.Add("Timestamp: $((Get-Date).ToString('yyyy-MM-ddTHH:mm:ss.fffK'))")
    $lines.Add("Diagnostic ID: $diagnosticId")
    $lines.Add("Level: $Level")
    $lines.Add("Context: $Context")
    $lines.Add("PowerShell: $($PSVersionTable.PSVersion)")
    $lines.Add("Run as: $([Security.Principal.WindowsIdentity]::GetCurrent().Name)")
    $lines.Add("Input CSV: $($InputCsv -join '; ')")
    $lines.Add("Output CSV: $OutputCsv")
    if ($Message) {
        $lines.Add("Summary: $Message")
    }
    if ($ErrorRecord) {
        $lines.Add((Format-ErrorRecord -ErrorRecord $ErrorRecord))
    }

    $diagnosticDirectory = Split-Path -Parent $script:DiagnosticLogPath
    if ($diagnosticDirectory -and -not (Test-Path -LiteralPath $diagnosticDirectory -PathType Container)) {
        New-Item -ItemType Directory -Path $diagnosticDirectory -Force | Out-Null
    }
    Add-Content -LiteralPath $script:DiagnosticLogPath -Value ($lines -join [Environment]::NewLine) -Encoding UTF8
    return $diagnosticId
}

try {
    $script:DiagnosticLogPath = [System.IO.Path]::GetFullPath($DiagnosticLog)
}
catch {
    throw "Diagnostic log path '$DiagnosticLog' is invalid: $($_.Exception.Message)"
}

function Invoke-StorageInventory {
$inputRows = [System.Collections.Generic.List[object]]::new()
$loadedFiles = [System.Collections.Generic.List[string]]::new()

foreach ($path in $InputCsv) {
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        $message = "Discovery CSV not found; skipping: $path"
        $diagnosticId = Write-Diagnostic -Level WARNING -Context 'Input validation' -Message $message
        Write-Warning "$message Diagnostic ID: $diagnosticId"
        continue
    }

    try {
        foreach ($row in @(Import-Csv -LiteralPath $path)) {
            $inputRows.Add($row)
        }
        $loadedFiles.Add((Resolve-Path -LiteralPath $path).Path)
    }
    catch {
        $diagnosticId = Write-Diagnostic -Level ERROR -Context "Importing discovery CSV '$path'" `
            -Message 'The discovery CSV could not be imported and was skipped.' -ErrorRecord $_
        Write-Warning "Could not import '$path'. Diagnostic ID: $diagnosticId. See '$script:DiagnosticLogPath'."
    }
}

if ($loadedFiles.Count -eq 0) {
    throw 'None of the supplied discovery CSV files could be found.'
}

$instances = @{}
$skippedRows = 0

foreach ($row in $inputRows) {
    $dataSource = Get-SqlDataSource -DiscoveryRow $row
    if (-not $dataSource) {
        $skippedRows++
        continue
    }

    $key = Get-InstanceKey -DiscoveryRow $row
    if (-not $instances.ContainsKey($key) -or
        ($dataSource -like 'tcp:*,*' -and $instances[$key].DataSource -notlike 'tcp:*,*')) {
        $instances[$key] = [pscustomobject]@{
            DiscoveryRow = $row
            DataSource   = $dataSource
        }
    }
}

if ($instances.Count -eq 0) {
    throw 'No connectable SQL Server instances were present in the discovery CSV files.'
}

$authentication = if ($Credential) {
    "SQL credential ($($Credential.UserName))"
}
else {
    "Current Windows user ($([Security.Principal.WindowsIdentity]::GetCurrent().Name))"
}

Write-Host ''
Write-Host '================================================================' -ForegroundColor Cyan
Write-Host ' SQL Server Database Storage Inventory' -ForegroundColor Cyan
Write-Host '================================================================' -ForegroundColor Cyan
Write-Host (" Input files       : {0}" -f $loadedFiles.Count)
Write-Host (" Unique instances  : {0}" -f $instances.Count)
Write-Host (" Authentication    : {0}" -f $authentication)
Write-Host (" System databases  : {0}" -f $(if ($ExcludeSystemDatabases) { 'Excluded' } else { 'Included' }))
Write-Host (" Output CSV        : {0}" -f $OutputCsv)
Write-Host (" Diagnostic log    : {0}" -f $script:DiagnosticLogPath)
if ($skippedRows -gt 0) {
    $message = "Skipped $skippedRows discovery row(s) without a connectable instance name."
    $diagnosticId = Write-Diagnostic -Level WARNING -Context 'Discovery row validation' -Message $message
    Write-Warning "$message Diagnostic ID: $diagnosticId"
}
Write-Host ''

$results = [System.Collections.Generic.List[object]]::new()
$current = 0

foreach ($instance in ($instances.Values | Sort-Object DataSource)) {
    $current++
    $row = $instance.DiscoveryRow
    $dataSource = $instance.DataSource
    $connection = $null
    $databaseCount = $null
    $databasesMeasured = 0
    $dataUsedMb = 0.0
    $logUsedMb = 0.0
    $issues = [System.Collections.Generic.List[string]]::new()
    $status = 'OK'

    Write-Progress -Activity 'Measuring SQL database storage' `
        -Status ("{0} ({1} of {2})" -f $dataSource, $current, $instances.Count) `
        -PercentComplete (($current / $instances.Count) * 100)

    try {
        $connectionString = New-SqlConnectionString -DataSource $dataSource -Database 'master'
        $connection = [System.Data.SqlClient.SqlConnection]::new($connectionString)
        $connection.Open()

        $databaseFilter = if ($ExcludeSystemDatabases) {
            "WHERE name NOT IN ('master', 'model', 'msdb', 'tempdb')"
        }
        else {
            ''
        }

        $databaseQuery = @"
SELECT name, CONVERT(varchar(20), DATABASEPROPERTYEX(name, 'Status')) AS DatabaseStatus
FROM master.dbo.sysdatabases
$databaseFilter
ORDER BY name;
"@
        $databases = Invoke-SqlDataTable -Connection $connection -Query $databaseQuery
        $databaseCount = $databases.Rows.Count

        $logUsageByDatabase = @{}
        try {
            $connection.ChangeDatabase('master')
            $logUsage = Invoke-SqlDataTable -Connection $connection -Query 'DBCC SQLPERF(LOGSPACE) WITH NO_INFOMSGS;'
            foreach ($logRow in $logUsage.Rows) {
                $databaseName = [string] $logRow.'Database Name'
                $logSizeMb = [double] $logRow.'Log Size (MB)'
                $logPercentUsed = [double] $logRow.'Log Space Used (%)'
                $logUsageByDatabase[$databaseName] = $logSizeMb * ($logPercentUsed / 100.0)
            }
        }
        catch {
            $diagnosticId = Write-Diagnostic -Level ERROR `
                -Context "Reading transaction-log usage on '$dataSource'" `
                -Message 'DBCC SQLPERF(LOGSPACE) failed.' -ErrorRecord $_
            $issues.Add("[Diagnostic $diagnosticId] Transaction-log usage unavailable: $($_.Exception.Message)")
            Write-Warning "Transaction-log usage failed on '$dataSource'. Diagnostic ID: $diagnosticId. See '$script:DiagnosticLogPath'."
        }

        foreach ($database in $databases.Rows) {
            $databaseName = [string] $database.name
            $databaseStatus = [string] $database.DatabaseStatus

            if ($databaseStatus -ne 'ONLINE') {
                $message = "Database '$databaseName' is $databaseStatus"
                $diagnosticId = Write-Diagnostic -Level WARNING `
                    -Context "Measuring database on '$dataSource'" -Message $message
                $issues.Add("[Diagnostic $diagnosticId] $message")
                continue
            }

            try {
                $connection.ChangeDatabase($databaseName)
                $dataUsage = Invoke-SqlDataTable -Connection $connection -Query @'
SELECT COALESCE(
    SUM(CONVERT(float, FILEPROPERTY(name, 'SpaceUsed'))) * 8.0 / 1024.0,
    0.0
) AS UsedDataMB
FROM dbo.sysfiles
WHERE groupid > 0;
'@
                $dataUsedMb += [double] $dataUsage.Rows[0].UsedDataMB
                if ($logUsageByDatabase.ContainsKey($databaseName)) {
                    $logUsedMb += [double] $logUsageByDatabase[$databaseName]
                }
                else {
                    $issues.Add("Transaction-log usage unavailable for database '$databaseName'")
                }
                $databasesMeasured++
            }
            catch {
                $diagnosticId = Write-Diagnostic -Level ERROR `
                    -Context "Measuring database '$databaseName' on '$dataSource'" `
                    -Message 'The database usage query failed.' -ErrorRecord $_
                $issues.Add("[Diagnostic $diagnosticId] Database '$databaseName': $($_.Exception.Message)")
                Write-Warning "Database '$databaseName' failed on '$dataSource'. Diagnostic ID: $diagnosticId. See '$script:DiagnosticLogPath'."
            }
        }

        if ($issues.Count -gt 0) {
            $status = 'PARTIAL'
        }
    }
    catch {
        $status = 'FAILED'
        $diagnosticId = Write-Diagnostic -Level ERROR -Context "Connecting to or inventorying '$dataSource'" `
            -Message 'The SQL Server instance could not be inventoried.' -ErrorRecord $_
        $issues.Add("[Diagnostic $diagnosticId] $($_.Exception.Message)")
        Write-Warning "Inventory failed for '$dataSource'. Diagnostic ID: $diagnosticId. See '$script:DiagnosticLogPath'."
    }
    finally {
        if ($connection) {
            try {
                $connection.Close()
                $connection.Dispose()
            }
            catch {
                $diagnosticId = Write-Diagnostic -Level ERROR -Context "Closing SQL connection to '$dataSource'" `
                    -Message 'The SQL connection could not be closed or disposed cleanly.' -ErrorRecord $_
                $issues.Add("[Diagnostic $diagnosticId] Connection cleanup failed: $($_.Exception.Message)")
                if ($status -eq 'OK') { $status = 'PARTIAL' }
                Write-Warning "Connection cleanup failed for '$dataSource'. Diagnostic ID: $diagnosticId. See '$script:DiagnosticLogPath'."
            }
        }
    }

    $totalUsedMb = $dataUsedMb + $logUsedMb
    $results.Add([pscustomobject][ordered]@{
        ServerName          = Get-PropertyValue -InputObject $row -Name 'ServerName'
        InstanceName        = Get-PropertyValue -InputObject $row -Name 'InstanceName'
        FQDN                = Get-PropertyValue -InputObject $row -Name 'FQDN'
        IPAddress           = Get-PropertyValue -InputObject $row -Name 'IPAddress'
        DataSource          = $dataSource
        SQLRelease          = Get-PropertyValue -InputObject $row -Name 'SQLRelease'
        SQLVersion          = Get-PropertyValue -InputObject $row -Name 'SQLVersion'
        DatabaseCount       = $databaseCount
        DatabasesMeasured   = $databasesMeasured
        DataUsedGB          = if ($status -eq 'FAILED') { $null } else { [Math]::Round($dataUsedMb / 1024.0, 3) }
        LogUsedGB           = if ($status -eq 'FAILED') { $null } else { [Math]::Round($logUsedMb / 1024.0, 3) }
        TotalStorageUsedGB  = if ($status -eq 'FAILED') { $null } else { [Math]::Round($totalUsedMb / 1024.0, 3) }
        Status              = $status
        ScanTime            = Get-Date -Format 's'
    })
}

Write-Progress -Activity 'Measuring SQL database storage' -Completed

$sorted = @($results | Sort-Object ServerName, InstanceName)
$sorted | Format-Table ServerName, InstanceName, DatabaseCount, DatabasesMeasured,
    DataUsedGB, LogUsedGB, TotalStorageUsedGB, Status -AutoSize

try {
    $outputDirectory = Split-Path -Parent $OutputCsv
    if ($outputDirectory -and -not (Test-Path -LiteralPath $outputDirectory -PathType Container)) {
        New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
    }

    $sorted | Export-Csv -LiteralPath $OutputCsv -NoTypeInformation -Encoding UTF8
    Write-Host ''
    Write-Host ("CSV exported to : {0}" -f $OutputCsv) -ForegroundColor Green
}
catch {
    throw
}

$okCount = @($sorted | Where-Object { $_.Status -eq 'OK' }).Count
$partialCount = @($sorted | Where-Object { $_.Status -eq 'PARTIAL' }).Count
$failedCount = @($sorted | Where-Object { $_.Status -eq 'FAILED' }).Count

Write-Host ("Successful        : {0}" -f $okCount)
Write-Host ("Partial           : {0}" -f $partialCount)
Write-Host ("Failed            : {0}" -f $failedCount)
Write-Host ("Elapsed           : {0:N1} seconds" -f ((Get-Date) - $script:StartTime).TotalSeconds)
Write-Host ''
}

try {
    Invoke-StorageInventory
}
catch {
    try {
        $diagnosticId = Write-Diagnostic -Level ERROR -Context 'Unhandled fatal script error' `
            -Message 'The inventory script terminated before successful completion.' -ErrorRecord $_
        Write-Error "Inventory terminated. Diagnostic ID: $diagnosticId. See '$script:DiagnosticLogPath'. $($_.Exception.Message)" `
            -ErrorAction Continue
    }
    catch {
        Write-Error "Inventory terminated and the diagnostic error could not be written: $($_.Exception.Message)" `
            -ErrorAction Continue
    }
    throw
}
