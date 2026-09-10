<#
    Site defaults for the OnpremMailboxMigration toolset.

    Copy this file to  OnPremMigration.Settings.psd1  (same folder) and edit the
    values. The real .psd1 is git-ignored. Every setting is optional and an
    explicit script parameter always overrides it.

    Consumed by: 01-Get-OnPremMigrationScope.ps1, 02-New-OnPremMoveRequest.ps1,
                 03/04/05/06 and the supporting scripts (for OutputRoot).
#>
@{
    # Parent folder for the per-wave output folders (<OutputRoot>\<WaveName>\).
    # Leave commented to write next to the scripts.
    # OutputRoot = 'D:\Migration\Waves'

    # 01 - readiness WARN thresholds and wave-split guardrails.
    MaxMailboxSizeGB = 50
    MaxWaveMailboxes = 500

    # 02 - default item limits for New-MoveRequest.
    DefaultBadItemLimit   = 0
    DefaultLargeItemLimit = 0
}
