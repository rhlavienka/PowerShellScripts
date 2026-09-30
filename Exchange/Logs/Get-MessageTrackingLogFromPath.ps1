<#
.SYNOPSIS
    Offline replacement for Get-MessageTrackingLog: searches Exchange message
    tracking log files (MSGTRK*.LOG) in a directory instead of querying a live
    Exchange server.

.DESCRIPTION
    Get-MessageTrackingLog can only read the logs of a running Exchange server
    (via the Exchange Management Shell). This script offers the same filters and
    returns objects with the same property names, but reads the raw log files
    from any directory - e.g. logs copied off a decommissioned server, a backup,
    or a support bundle. No Exchange module is needed.

    How it works:
      - Every file matching -Filter (default MSGTRK*.LOG, which covers MSGTRK,
        MSGTRKMA, MSGTRKMD and MSGTRKMS) in -Path is read, oldest first.
      - The column layout is taken from each file's "#Fields:" header, so logs
        from Exchange 2010 through Exchange SE are handled.
      - Files are opened with shared read/write access, so logs the Transport
        service is still writing to can be read too.
      - Files that cannot contain entries between -Start and -End are skipped
        without being opened. The time range of a file is derived from its name
        (MSGTRK<yyyyMMddHH>-<n>.LOG) and the name of the next file of the same
        series, so this also works on logs copied to another machine.
      - Each line is checked with a quick text match before the full CSV parse,
        so narrow searches through large log sets are fast.

    Filter semantics follow Get-MessageTrackingLog:
      - All given filters must match (AND).
      - -Recipients matches if any recipient of the entry equals any of the
        given addresses.
      - -MessageSubject is a case-insensitive substring match.
      - Sender, recipient and ID filters are case-insensitive exact matches;
        angle brackets around -MessageId and -Reference are optional.
      - -Start / -End are in local time (log timestamps are UTC and are
        converted to local time in the output, like the cmdlet does).
      - -ResultSize defaults to 1000; a warning is shown when more entries
        match. Use -ResultSize Unlimited to return everything.

    Differences from Get-MessageTrackingLog:
      - -Path (required) replaces -Server / -DomainController.
      - -EventId and -Source accept several values.
      - Properties are in the cmdlet's order; the log columns the cmdlet does
        not show (OriginalServerIp, LogId) and LogFile (the file the entry
        came from) are appended at the end. There is no RunspaceId.

.PARAMETER Path
    Directory (or several) containing the message tracking log files. On a live
    server the default location is
    %ExchangeInstallPath%TransportRoles\Logs\MessageTracking.

.PARAMETER Recurse
    Also search subdirectories of -Path (e.g. one subfolder per server).

.PARAMETER Filter
    File name filter. Default: MSGTRK*.LOG.

.PARAMETER Start
    Return only entries logged at or after this local date/time.

.PARAMETER End
    Return only entries logged at or before this local date/time.

.PARAMETER EventId
    Event ID(s) to return, e.g. RECEIVE, SEND, DELIVER, FAIL, DEFER, DSN,
    EXPAND, REDIRECT, RESOLVE, SUBMIT, TRANSFER, HAREDIRECT, BADMAIL.

.PARAMETER Source
    Source(s) to return, e.g. SMTP, STOREDRIVER, ROUTING, AGENT, DSN, MAILBOXRULE.

.PARAMETER Sender
    Sender SMTP address (exact match).

.PARAMETER Recipients
    Recipient SMTP address(es). An entry matches if any of its recipients is
    one of these addresses.

.PARAMETER MessageSubject
    Text that must appear in the message subject (substring match).

.PARAMETER MessageId
    Internet Message-ID (the Message-ID header). Angle brackets are optional.

.PARAMETER InternalMessageId
    Exchange internal message ID (per-server number).

.PARAMETER NetworkMessageId
    Network message ID (GUID shared by all copies of a message across servers).

.PARAMETER Reference
    Value that must appear in the reference field (e.g. the Message-ID of the
    original message on DSN entries). Angle brackets are optional.

.PARAMETER TransportTrafficType
    Transport traffic type (e.g. Email, Journal). Exact match.

.PARAMETER ResultSize
    Maximum number of entries to return: a number, or Unlimited. Default: 1000.

.PARAMETER ScanAllFiles
    Read every matching file, even ones whose name says it is outside
    -Start / -End (e.g. when the log files were renamed).

.EXAMPLE
    .\Get-MessageTrackingLogFromPath.ps1 -Path D:\Logs\MessageTracking -Sender john@contoso.com -Start (Get-Date).AddDays(-2)
    Everything sent by john@contoso.com in the last two days.

.EXAMPLE
    .\Get-MessageTrackingLogFromPath.ps1 -Path D:\Logs -Recurse -Recipients anna@contoso.com -EventId DELIVER,FAIL -ResultSize Unlimited
    All deliveries and failures to anna@contoso.com, across per-server subfolders.

.EXAMPLE
    .\Get-MessageTrackingLogFromPath.ps1 -Path D:\Logs\MessageTracking -MessageId '<abc123@contoso.com>' |
        Sort-Object Timestamp | Format-Table Timestamp,EventId,Source,ServerHostname,Recipients,RecipientStatus -AutoSize
    Follows one message through all its tracking events.

.EXAMPLE
    .\Get-MessageTrackingLogFromPath.ps1 -Path D:\Logs\MessageTracking -EventId RECEIVE -Source SMTP -ResultSize Unlimited |
        Group-Object ClientIp | Sort-Object Count -Descending | Select-Object Count,Name
    Which client IPs submitted mail over SMTP, and how often.

.EXAMPLE
    .\Get-MessageTrackingLogFromPath.ps1 -Path D:\Logs\MessageTracking -MessageSubject 'Invoice' -ResultSize Unlimited |
        Select-Object Timestamp,EventId,Sender,@{n='Recipients';e={$_.Recipients -join ';'}},MessageSubject |
        Export-Csv .\invoice.csv -NoTypeInformation -Encoding UTF8
    Exports matching entries to CSV (array properties joined with ';').

.NOTES
    Requires Windows PowerShell 5.1 or PowerShell 7+. No Exchange module needed.

    Version: 1.0 (2026-09-30)
    Author:  Richard Hlavienka (richard.hlavienka@elyvyn.com)

    Changelog:
      1.0 (2026-09-30) - Initial version
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Container })]
    [string[]]$Path,

    [switch]$Recurse,

    [string]$Filter = 'MSGTRK*.LOG',

    [datetime]$Start,

    [datetime]$End,

    [string[]]$EventId,

    [string[]]$Source,

    [string]$Sender,

    [string[]]$Recipients,

    [string]$MessageSubject,

    [string]$MessageId,

    [string]$InternalMessageId,

    [string]$NetworkMessageId,

    [string]$Reference,

    [string]$TransportTrafficType,

    [ValidateScript({ $_ -eq 'Unlimited' -or ($_ -as [int]) -gt 0 })]
    [string]$ResultSize = '1000',

    [switch]$ScanAllFiles
)

# --- Log reader (compiled once; per-line work in script would be far too slow) --
if (-not ('MsgTrkLogReader' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Management.Automation;
using System.Text;

// Search filters; empty members are ignored. Call Prepare() after setting them.
public sealed class MsgTrkLogFilter {
    public string StartText, EndText;   // UTC, "yyyy-MM-ddTHH:mm:ss.fff[Z]"
    public string Sender, MessageId, InternalMessageId, NetworkMessageId, Reference, MessageSubject, TransportTrafficType;
    public string[] Recipients, EventIds, Sources;

    internal Dictionary<string, bool> RecipientSet, EventSet, SourceSet;
    internal List<string[]> Needles;    // text that must occur in the raw line (any of each group)

    public void Prepare() {
        Sender = Norm(Sender); InternalMessageId = Norm(InternalMessageId); NetworkMessageId = Norm(NetworkMessageId);
        TransportTrafficType = Norm(TransportTrafficType);
        MessageId = MsgTrkLogReader.TrimAngles(Norm(MessageId)); Reference = MsgTrkLogReader.TrimAngles(Norm(Reference));
        if (string.IsNullOrEmpty(MessageSubject)) { MessageSubject = null; }
        RecipientSet = ToSet(Recipients); EventSet = ToSet(EventIds); SourceSet = ToSet(Sources);

        Needles = new List<string[]>();
        foreach (string s in new[] { Sender, InternalMessageId, NetworkMessageId, TransportTrafficType }) { if (s != null) { Needles.Add(new[] { s }); } }
        foreach (string s in new[] { MessageId, Reference, MessageSubject }) { if (s != null) { Needles.Add(new[] { s.Replace("\"", "\"\"") }); } }
        foreach (var set in new[] { RecipientSet, EventSet, SourceSet }) { if (set != null) { Needles.Add(new List<string>(set.Keys).ToArray()); } }
    }

    static string Norm(string s) { if (s == null) { return null; } s = s.Trim(); return s.Length == 0 ? null : s; }

    static Dictionary<string, bool> ToSet(string[] values) {
        if (values == null) { return null; }
        var set = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
        foreach (string v in values) { string n = Norm(v); if (n != null) { set[n] = true; } }
        return set.Count == 0 ? null : set;
    }
}

// Plain wrapper, because PowerShell cannot call the explicit interface
// members of a compiler-generated iterator.
public sealed class MsgTrkLogCursor : IDisposable {
    readonly IEnumerator<PSObject> inner;
    internal MsgTrkLogCursor(IEnumerator<PSObject> inner) { this.inner = inner; }
    public bool MoveNext() { return inner.MoveNext(); }
    public PSObject Current { get { return inner.Current; } }
    public void Dispose() { inner.Dispose(); }
}

public static class MsgTrkLogReader {
    public const string TypeName = "MessageTrackingLogFileEntry";
    const StringComparison IC = StringComparison.OrdinalIgnoreCase;

    // Log column -> property name used by Get-MessageTrackingLog
    static readonly Dictionary<string, string> PropertyMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) {
        {"date-time","Timestamp"}, {"client-ip","ClientIp"}, {"client-hostname","ClientHostname"},
        {"server-ip","ServerIp"}, {"server-hostname","ServerHostname"}, {"source-context","SourceContext"},
        {"connector-id","ConnectorId"}, {"source","Source"}, {"event-id","EventId"},
        {"internal-message-id","InternalMessageId"}, {"message-id","MessageId"}, {"network-message-id","NetworkMessageId"},
        {"recipient-address","Recipients"}, {"recipient-status","RecipientStatus"}, {"total-bytes","TotalBytes"},
        {"recipient-count","RecipientCount"}, {"related-recipient-address","RelatedRecipientAddress"}, {"reference","Reference"},
        {"message-subject","MessageSubject"}, {"sender-address","Sender"}, {"return-path","ReturnPath"},
        {"message-info","MessageInfo"}, {"directionality","Directionality"}, {"tenant-id","TenantId"},
        {"original-client-ip","OriginalClientIp"}, {"original-server-ip","OriginalServerIp"}, {"custom-data","EventData"},
        {"transport-traffic-type","TransportTrafficType"}, {"log-id","LogId"}, {"schema-version","SchemaVersion"}
    };

    // Property order of Get-MessageTrackingLog; other log columns follow at the end
    static readonly string[] CmdletOrder = {
        "Timestamp", "ClientIp", "ClientHostname", "ServerIp", "ServerHostname", "SourceContext", "ConnectorId", "Source",
        "EventId", "InternalMessageId", "MessageId", "NetworkMessageId", "Recipients", "RecipientStatus", "TotalBytes",
        "RecipientCount", "RelatedRecipientAddress", "Reference", "MessageSubject", "Sender", "ReturnPath", "Directionality",
        "TenantId", "OriginalClientIp", "MessageInfo", "MessageLatency", "MessageLatencyType", "EventData",
        "TransportTrafficType", "SchemaVersion"
    };

    // Used when a file has no "#Fields:" header (Exchange 2013+ layout)
    const string DefaultFields = "date-time,client-ip,client-hostname,server-ip,server-hostname,source-context,connector-id,source,event-id,internal-message-id,message-id,network-message-id,recipient-address,recipient-status,total-bytes,recipient-count,related-recipient-address,reference,message-subject,sender-address,return-path,message-info,directionality,tenant-id,original-client-ip,original-server-ip,custom-data,transport-traffic-type,log-id,schema-version";

    // Opens one log file; the cursor returns the matching entries in file order.
    public static MsgTrkLogCursor Open(string path, MsgTrkLogFilter filter) {
        return new MsgTrkLogCursor(ReadFile(path, filter).GetEnumerator());
    }

    static IEnumerable<PSObject> ReadFile(string path, MsgTrkLogFilter filter) {
        string[] props = ToProperties(DefaultFields);
        using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
        using (var reader = new StreamReader(stream, Encoding.UTF8, true)) {
            string line;
            while ((line = reader.ReadLine()) != null) {
                if (line.Length == 0) { continue; }
                if (line[0] == '#') {
                    if (line.StartsWith("#Fields:", IC)) { props = ToProperties(line.Substring(8)); }
                    continue;
                }
                if (!PreMatch(line, filter)) { continue; }

                string[] values = Split(line);
                var raw = new Dictionary<string, string>(props.Length, StringComparer.OrdinalIgnoreCase);
                for (int i = 0; i < props.Length; i++) { raw[props[i]] = i < values.Length ? values[i] : ""; }

                if (Match(raw, filter)) { yield return ToObject(raw, props, path); }
            }
        }
    }

    static string[] ToProperties(string fields) {
        string[] names = fields.Trim().Split(',');
        for (int i = 0; i < names.Length; i++) {
            string f = names[i].Trim();
            string p;
            if (!PropertyMap.TryGetValue(f, out p)) {
                var sb = new StringBuilder();
                foreach (string part in f.Split('-')) { if (part.Length > 0) { sb.Append(char.ToUpperInvariant(part[0])).Append(part.Substring(1)); } }
                p = sb.ToString();
            }
            names[i] = p;
        }
        return names;
    }

    // Cheap checks on the raw line before the full CSV parse
    static bool PreMatch(string line, MsgTrkLogFilter f) {
        if (f.StartText != null || f.EndText != null) {
            // Log timestamps are fixed-width ISO 8601 UTC, so they compare as strings
            int comma = line.IndexOf(',');
            string ts = comma > 0 ? line.Substring(0, comma) : line;
            if (f.StartText != null && string.CompareOrdinal(ts, f.StartText) < 0) { return false; }
            if (f.EndText != null && string.CompareOrdinal(ts, f.EndText) > 0) { return false; }
        }
        foreach (string[] group in f.Needles) {
            bool hit = false;
            foreach (string needle in group) { if (line.IndexOf(needle, IC) >= 0) { hit = true; break; } }
            if (!hit) { return false; }
        }
        return true;
    }

    // Exact filter checks, same semantics as Get-MessageTrackingLog
    static bool Match(Dictionary<string, string> raw, MsgTrkLogFilter f) {
        if (f.Sender != null && !string.Equals(Get(raw, "Sender"), f.Sender, IC)) { return false; }
        if (f.EventSet != null && !f.EventSet.ContainsKey(Get(raw, "EventId"))) { return false; }
        if (f.SourceSet != null && !f.SourceSet.ContainsKey(Get(raw, "Source"))) { return false; }
        if (f.MessageId != null && !string.Equals(TrimAngles(Get(raw, "MessageId")), f.MessageId, IC)) { return false; }
        if (f.InternalMessageId != null && !string.Equals(Get(raw, "InternalMessageId"), f.InternalMessageId, StringComparison.Ordinal)) { return false; }
        if (f.NetworkMessageId != null && !string.Equals(Get(raw, "NetworkMessageId"), f.NetworkMessageId, IC)) { return false; }
        if (f.TransportTrafficType != null && !string.Equals(Get(raw, "TransportTrafficType"), f.TransportTrafficType, IC)) { return false; }
        if (f.MessageSubject != null && Get(raw, "MessageSubject").IndexOf(f.MessageSubject, IC) < 0) { return false; }
        if (f.RecipientSet != null) {
            bool hit = false;
            foreach (string r in SplitList(Get(raw, "Recipients"))) { if (f.RecipientSet.ContainsKey(r)) { hit = true; break; } }
            if (!hit) { return false; }
        }
        if (f.Reference != null) {
            bool hit = false;
            foreach (string r in SplitList(Get(raw, "Reference"))) { if (string.Equals(TrimAngles(r), f.Reference, IC)) { hit = true; break; } }
            if (!hit) { return false; }
        }
        return true;
    }

    // Builds the output object with the same property names and types as the cmdlet
    static PSObject ToObject(Dictionary<string, string> raw, string[] props, string path) {
        var values = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
        foreach (string p in props) {
            string v = raw[p];
            switch (p) {
                case "Timestamp":
                    DateTime dt;
                    values[p] = DateTime.TryParse(v, CultureInfo.InvariantCulture, DateTimeStyles.AdjustToUniversal, out dt) ? (object)dt.ToLocalTime() : v;
                    break;
                case "Recipients":
                case "RecipientStatus":
                    values[p] = SplitList(v);
                    break;
                case "Reference":
                    string[] refs = SplitList(v);
                    values[p] = refs.Length == 0 ? null : refs;
                    break;
                case "TotalBytes":
                case "RecipientCount":
                    long n;
                    values[p] = long.TryParse(v, NumberStyles.Integer, CultureInfo.InvariantCulture, out n) ? (object)n : null;
                    break;
                case "EventData":
                    values[p] = ParseEventData(v);
                    break;
                case "MessageInfo":
                case "SourceContext":
                    values[p] = Unquote(v);
                    break;
                default:
                    values[p] = v;
                    break;
            }
        }

        // End-to-end latency comes from the E2ELatency event data (seconds), like the cmdlet
        values["MessageLatency"] = null;
        values["MessageLatencyType"] = "None";
        object ed;
        if (values.TryGetValue("EventData", out ed)) {
            foreach (var kv in (KeyValuePair<string, object>[])ed) {
                double sec;
                if (kv.Key == "E2ELatency" && double.TryParse(kv.Value as string, NumberStyles.Float, CultureInfo.InvariantCulture, out sec)) {
                    // Rounded to whole ticks, so .NET Framework and .NET give the same value
                    values["MessageLatency"] = TimeSpan.FromTicks((long)Math.Round(sec * TimeSpan.TicksPerSecond));
                    values["MessageLatencyType"] = "EndToEnd";
                    break;
                }
            }
        }

        var obj = new PSObject();
        obj.TypeNames.Insert(0, TypeName);
        foreach (string p in CmdletOrder) {
            object v;
            if (values.TryGetValue(p, out v)) { obj.Properties.Add(new PSNoteProperty(p, v)); values.Remove(p); }
        }
        foreach (string p in props) {
            object v;
            if (values.TryGetValue(p, out v)) { obj.Properties.Add(new PSNoteProperty(p, v)); }
        }
        obj.Properties.Add(new PSNoteProperty("LogFile", path));
        return obj;
    }

    static string Get(Dictionary<string, string> raw, string key) {
        string v;
        return raw.TryGetValue(key, out v) ? v : "";
    }

    public static string TrimAngles(string s) {
        return s == null ? null : s.Trim().Trim('<', '>');
    }

    // Splits one CSV line; fields may be in double quotes ("" = literal quote)
    public static string[] Split(string line) {
        var fields = new List<string>(32);
        var sb = new StringBuilder(line.Length);
        bool quoted = false;
        for (int i = 0; i < line.Length; i++) {
            char c = line[i];
            if (quoted) {
                if (c == '"') {
                    if (i + 1 < line.Length && line[i + 1] == '"') { sb.Append('"'); i++; }
                    else { quoted = false; }
                } else { sb.Append(c); }
            } else if (c == '"') { quoted = true; }
            else if (c == ',') { fields.Add(sb.ToString()); sb.Length = 0; }
            else { sb.Append(c); }
        }
        fields.Add(sb.ToString());
        return fields.ToArray();
    }

    // Splits a ';'-separated list. Items that contain ';' themselves are
    // written by Exchange in single quotes: a;'b; c';d -> a | b; c | d
    public static string[] SplitList(string value) {
        var items = new List<string>();
        if (string.IsNullOrEmpty(value)) { return items.ToArray(); }
        int pos = 0;
        while (pos < value.Length) {
            int end;
            if (value[pos] == '\'') {
                end = value.IndexOf("';", pos + 1, StringComparison.Ordinal);
                if (end < 0 && value[value.Length - 1] == '\'' && value.Length - 1 > pos) { end = value.Length - 1; }
                if (end >= 0) {
                    items.Add(value.Substring(pos + 1, end - pos - 1));
                    pos = end + 2;
                    continue;
                }
            }
            end = value.IndexOf(';', pos);
            if (end < 0) { end = value.Length; }
            if (end > pos) { items.Add(value.Substring(pos, end - pos)); }
            pos = end + 1;
        }
        return items.ToArray();
    }

    // "S:Key1=Value1;'S:Key2=a;b'" -> key/value pairs (type prefix dropped, like the cmdlet)
    public static KeyValuePair<string, object>[] ParseEventData(string value) {
        var pairs = new List<KeyValuePair<string, object>>();
        foreach (string entry in SplitList(value)) {
            string item = entry;
            int colon = item.IndexOf(':');
            int eq = item.IndexOf('=');
            if (colon > 0 && colon <= 3 && (eq < 0 || colon < eq)) { item = item.Substring(colon + 1); eq = item.IndexOf('='); }
            if (eq < 0) { pairs.Add(new KeyValuePair<string, object>(item, null)); }
            else { pairs.Add(new KeyValuePair<string, object>(item.Substring(0, eq), item.Substring(eq + 1))); }
        }
        return pairs.ToArray();
    }

    // Removes the single quotes Exchange puts around a value that contains ';'
    public static string Unquote(string value) {
        if (value != null && value.Length >= 2 && value[0] == '\'' && value[value.Length - 1] == '\'' && value.IndexOf(';') >= 0) {
            return value.Substring(1, value.Length - 2);
        }
        return value;
    }
}
'@
}

# --- Output type: same default columns as Get-MessageTrackingLog -------------
Update-TypeData -TypeName ([MsgTrkLogReader]::TypeName) -DefaultDisplayPropertySet Timestamp, EventId, Source, Sender, Recipients, MessageSubject -Force

# --- Filters ------------------------------------------------------------------
$limit = if ($ResultSize -eq 'Unlimited') { [int]::MaxValue } else { [int]$ResultSize }
$startUtc = if ($PSBoundParameters.ContainsKey('Start')) { $Start.ToUniversalTime() } else { $null }
$endUtc = if ($PSBoundParameters.ContainsKey('End')) { $End.ToUniversalTime() } else { $null }

$searchFilter = [MsgTrkLogFilter]::new()
if ($startUtc) { $searchFilter.StartText = $startUtc.ToString('yyyy-MM-ddTHH:mm:ss.fff') }
if ($endUtc) { $searchFilter.EndText = $endUtc.ToString('yyyy-MM-ddTHH:mm:ss.fff') + 'Z' }
$searchFilter.Sender = $Sender
$searchFilter.Recipients = $Recipients
$searchFilter.EventIds = $EventId
$searchFilter.Sources = $Source
$searchFilter.MessageSubject = $MessageSubject
$searchFilter.MessageId = $MessageId
$searchFilter.InternalMessageId = $InternalMessageId
$searchFilter.NetworkMessageId = $NetworkMessageId
$searchFilter.Reference = $Reference
$searchFilter.TransportTrafficType = $TransportTrafficType
$searchFilter.Prepare()

# --- Collect log files, oldest first -----------------------------------------
# File names are MSGTRK[MA|MD|MS]<yyyyMMdd[HH]>-<n>.LOG; the timestamp in the name
# is the UTC hour (or day, on old versions) in which the file was started.
$logFiles = foreach ($dir in $Path) {
    foreach ($f in (Get-ChildItem -LiteralPath $dir -Filter $Filter -File -Recurse:$Recurse -ErrorAction SilentlyContinue)) {
        $started = $f.CreationTimeUtc; $slot = $null; $series = $f.FullName
        if ($f.Name -match '^(MSGTRK[A-Z]*)(\d{8})(\d{2})?-\d+\.LOG$') {
            $hour = if ($Matches[3]) { $Matches[3] } else { '00' }
            $slot = if ($Matches[3]) { [timespan]::FromHours(1) } else { [timespan]::FromDays(1) }
            $series = Join-Path $f.DirectoryName $Matches[1]
            $started = [datetime]::ParseExact($Matches[2] + $hour, 'yyyyMMddHH', [cultureinfo]::InvariantCulture,
                [System.Globalization.DateTimeStyles]::AssumeUniversal -bor [System.Globalization.DateTimeStyles]::AdjustToUniversal)
        }
        [pscustomobject]@{ File = $f; Series = $series; Started = $started; Slot = $slot; EndsBefore = $null }
    }
}
$logFiles = @($logFiles | Sort-Object Started, { $_.File.Name })

if ($logFiles.Count -eq 0) {
    Write-Warning "No log files matching '$Filter' found in: $($Path -join ', ')"
    return
}

if (-not $ScanAllFiles) {
    # A file holds no entries newer than the start of the next file of the same
    # series (MSGTRK, MSGTRKMD, ...), which began before the end of its name slot.
    # Unlike LastWriteTime, this survives copying the logs to another machine.
    foreach ($group in ($logFiles | Where-Object Slot | Group-Object Series)) {
        $list = @($group.Group)
        for ($i = 0; $i -lt $list.Count - 1; $i++) { $list[$i].EndsBefore = $list[$i + 1].Started + $list[$i + 1].Slot }
    }
    $logFiles = @($logFiles | Where-Object {
            -not ($endUtc -and $_.Started -gt $endUtc) -and
            -not ($startUtc -and $_.File.LastWriteTimeUtc -lt $startUtc) -and
            -not ($startUtc -and $_.EndsBefore -and $_.EndsBefore -le $startUtc)
        })
    if ($logFiles.Count -eq 0) {
        Write-Warning "No log file covers the -Start / -End time range. Use -ScanAllFiles to read all files anyway."
        return
    }
}
$files = @($logFiles | ForEach-Object File)
Write-Verbose "Searching $($files.Count) log file(s)."

# --- Scan ---------------------------------------------------------------------
$returned = 0
$moreAvailable = $false
$fileIndex = 0

:files foreach ($file in $files) {
    $fileIndex++
    Write-Progress -Activity 'Searching message tracking logs' -Status "$($file.Name) ($fileIndex / $($files.Count)) - $returned match(es)" -PercentComplete (100 * $fileIndex / $files.Count)

    $entries = $null
    try {
        $entries = [MsgTrkLogReader]::Open($file.FullName, $searchFilter)
        while ($entries.MoveNext()) {
            if ($returned -ge $limit) { $moreAvailable = $true; break files }
            $entries.Current
            $returned++
        }
    }
    catch {
        Write-Warning "Cannot read '$($file.FullName)': $($_.Exception.GetBaseException().Message)"
    }
    finally {
        # Closes the file also when the scan stops early
        if ($entries) { $entries.Dispose() }
    }
}

Write-Progress -Activity 'Searching message tracking logs' -Completed

if ($moreAvailable) {
    Write-Warning "There are more results than the $limit returned. Use -ResultSize with a larger value or 'Unlimited' to see them all."
}
