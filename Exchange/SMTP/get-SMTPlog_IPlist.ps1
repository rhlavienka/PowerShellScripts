<#
.SYNOPSIS
    Analyzes Exchange SMTP protocol logs to report relay usage by client IP and receive connector.

.DESCRIPTION
    Parses the most recent Exchange FrontEnd SMTP Receive protocol logs
    (FrontEnd\ProtocolLog\SmtpReceive) and counts sent messages (MAIL FROM),
    grouped by client source IP address and Receive Connector.
    Results are exported to CSV, opened in Out-GridView, and printed to the console.

    The script has no parameters - configuration (log path, number of logs to
    process, output file) is set via the variables at the top of the script
    ($LogPath, $LastLogs, $CsvFile). Edit them directly if your environment
    differs from the defaults.

.EXAMPLE
    .\get-SMTPlog_IPlist.ps1
    Processes the 20 most recent SMTP Receive protocol logs on the local Exchange
    server, exports a CSV to C:\Temp, and opens the results in Out-GridView.

.NOTES
    Version: 1.0 (2026-08-20)
    Author:  Richard Hlavienka (richard.hlavienka@elyvyn.com)
    Run locally on an Exchange server with access to:
      C:\Program Files\Microsoft\Exchange Server\V15\TransportRoles\Logs\FrontEnd\ProtocolLog\SmtpReceive
#>

# Path to SMTP Receive logs
$LogPath = "C:\Program Files\Microsoft\Exchange Server\V15\TransportRoles\Logs\FrontEnd\ProtocolLog\SmtpReceive"

# Number of most recent logs to process
$LastLogs = 20

# Output file
$TimeStamp = Get-Date -Format "yyyyMMdd_HHmmss"
$CsvFile = "C:\Temp\SMTP_Relay_Usage_$TimeStamp.csv"

Write-Host ""
Write-Host "Loading SMTP Protocol Logs..." -ForegroundColor Cyan

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
            Connector      = $Record.'connector-id'
            RemoteEndpoint = $Record.'remote-endpoint'
            Event          = $Record.event
            Data           = $Record.data
        }
    }
}

Write-Host ""
Write-Host "Analyzing MAIL FROM records..." -ForegroundColor Cyan

$Results = $AllRecords |
    Where-Object {

        # Actual message send
        $_.Data -match 'MAIL FROM:'
    } |
    ForEach-Object {

        $IP = $null

        if ($_.RemoteEndpoint)
        {
            $IP = ($_.RemoteEndpoint -split ':')[0]
        }

        [PSCustomObject]@{
            ClientIP  = $IP
            Connector = $_.Connector
        }
    } |
    Where-Object {
        $_.ClientIP
    } |
    Group-Object ClientIP,Connector |
    Sort-Object Count -Descending |
    Select-Object @{
            Name = "Messages"
            Expression = { $_.Count }
        },
        @{
            Name = "ClientIP"
            Expression = { $_.Group[0].ClientIP }
        },
        @{
            Name = "Connector"
            Expression = { $_.Group[0].Connector }
        }

# Export CSV
$Results |
    Export-Csv `
        -Path $CsvFile `
        -NoTypeInformation `
        -Encoding UTF8

# GridView
$Results |
    Out-GridView -Title "SMTP Relay Usage"

# Console
$Results | Format-Table -AutoSize

Write-Host ""
Write-Host "Results found : $($Results.Count)" -ForegroundColor Green
Write-Host "CSV exported  : $CsvFile" -ForegroundColor Green
