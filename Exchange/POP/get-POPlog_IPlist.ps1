<#
.SYNOPSIS
    Analyzes Exchange POP3 protocol logs to report client logon usage by IP address.

.DESCRIPTION
    Parses the most recent Exchange POP3 protocol logs and counts client USER /
    APOP logon commands, grouped by client source IP address and the server
    endpoint (IP:Port) the client connected to (useful to tell FrontEnd from
    Backend/MBX connections apart).
    Results are exported to CSV, opened in Out-GridView, and printed to the console.

    The script has no parameters - configuration (log path, number of logs to
    process, output file) is set via the variables at the top of the script
    ($LogPath, $LastLogs, $CsvFile). Edit them directly if your environment
    differs from the defaults.

    POP3 protocol logging is disabled by default on Exchange Server and must be
    enabled before this script has any data to analyze. Run on the Exchange
    server (or against every Mailbox server if load balanced) as Administrator:

        Set-PopSettings -ProtocolLogEnabled $true `
            -LogFileLocation "C:\Program Files\Microsoft\Exchange Server\V15\Logging\Pop3"
        Restart-Service MSExchangePOP3
        Restart-Service MSExchangePOP3BE

.EXAMPLE
    .\get-POPlog_IPlist.ps1
    Processes the 20 most recent POP3 protocol logs on the local Exchange
    server, exports a CSV to C:\Temp, and opens the results in Out-GridView.

.NOTES
    Version: 1.0 (2026-08-21)
    Author:  Richard Hlavienka (richard.hlavienka@elyvyn.com)
    Run locally on an Exchange server with access to:
      C:\Program Files\Microsoft\Exchange Server\V15\Logging\Pop3
#>

# Path to POP3 protocol logs
$LogPath = "C:\Program Files\Microsoft\Exchange Server\V15\Logging\Pop3"

# Number of most recent logs to process
$LastLogs = 20

# Output file
$TimeStamp = Get-Date -Format "yyyyMMdd_HHmmss"
$CsvFile = "C:\Temp\POP3_Logon_Usage_$TimeStamp.csv"

Write-Host ""
Write-Host "Loading POP3 Protocol Logs..." -ForegroundColor Cyan

$AllRecords = foreach ($File in (Get-ChildItem "$LogPath\*.log" |
                                 Sort-Object LastWriteTime |
                                 Select-Object -Last $LastLogs))
{
    Write-Host "Processing: $($File.Name)" -ForegroundColor Yellow

    $Lines = Get-Content $File.FullName

    $FieldsLine = $Lines |
        Where-Object { $_ -like '#Fields:*' } |
        Select-Object -First 1

    if (-not $FieldsLine)
    {
        Write-Warning "Skipping $($File.Name) - no #Fields header found"
        continue
    }

    $Headers = ($FieldsLine -replace '^#Fields:\s*','') -split ','

    $DataLines = $Lines |
        Where-Object {
            $_ -and
            $_ -notmatch '^#'
        }

    $Records = $DataLines |
        ConvertFrom-Csv -Header $Headers

    foreach ($Record in $Records)
    {
        [PSCustomObject]@{
            ClientIP   = $Record.ClientIP
            ServerIP   = $Record.ServerIP
            ServerPort = $Record.ServerPort
            Event      = $Record.event
            Data       = $Record.data
        }
    }
}

Write-Host ""
Write-Host "Analyzing USER / APOP records..." -ForegroundColor Cyan

$Results = $AllRecords |
    Where-Object {

        # Client logon commands (plaintext or APOP)
        $_.Data -match '(?i)^USER\s|^APOP\s'
    } |
    Where-Object {
        $_.ClientIP
    } |
    Group-Object ClientIP,ServerIP,ServerPort |
    Sort-Object Count -Descending |
    Select-Object @{
            Name = "Logons"
            Expression = { $_.Count }
        },
        @{
            Name = "ClientIP"
            Expression = { $_.Group[0].ClientIP }
        },
        @{
            Name = "ServerEndpoint"
            Expression = { "$($_.Group[0].ServerIP):$($_.Group[0].ServerPort)" }
        }

# Export CSV
$Results |
    Export-Csv `
        -Path $CsvFile `
        -NoTypeInformation `
        -Encoding UTF8

# GridView
$Results |
    Out-GridView -Title "POP3 Logon Usage"

# Console
$Results | Format-Table -AutoSize

Write-Host ""
Write-Host "Results found : $($Results.Count)" -ForegroundColor Green
Write-Host "CSV exported  : $CsvFile" -ForegroundColor Green
