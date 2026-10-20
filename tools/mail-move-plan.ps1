# Read-only, versioned QA adapter. No Graph token, no provider mutation.
[CmdletBinding()]
param(
    [switch]$Fixture,
    [string]$SnapshotJson,
    [string]$MailboxGuid,
    [string]$ImmutableId,
    [string]$DestinationKey
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Require([bool]$Ok, [string]$Code) {
    if (-not $Ok) { throw "PRECONDITION_FAILED:$Code" }
}

if ($Fixture) {
    $MailboxGuid = '11111111-1111-4111-8111-111111111111'
    $ImmutableId = 'AQMk-SYNTHETIC-IMMUTABLE-001'
    $DestinationKey = 'Institutions/DaZ'
    $SnapshotJson = @'
{"schema":"mail-move-snapshot/v1","identity":{"mailboxGuid":"11111111-1111-4111-8111-111111111111","idType":"ImmutableId"},"message":{"id":"AQMk-SYNTHETIC-IMMUTABLE-001","parentFolderId":"folder-synthetic-inbox","changeKey":"ck-synthetic-1","isDraft":false,"flag":{"flagStatus":"notFlagged"},"categories":[]},"folders":[{"key":"Institutions/DaZ","id":"folder-synthetic-daz","mailboxGuid":"11111111-1111-4111-8111-111111111111"}]}
'@
}

Require (-not [string]::IsNullOrWhiteSpace($SnapshotJson)) 'snapshot_required'
Require (-not [string]::IsNullOrWhiteSpace($MailboxGuid)) 'mailbox_required'
Require (-not [string]::IsNullOrWhiteSpace($ImmutableId)) 'immutable_id_required'
Require (-not [string]::IsNullOrWhiteSpace($DestinationKey)) 'destination_required'
$guidValue = [guid]::Empty
Require ([guid]::TryParse($MailboxGuid, [ref]$guidValue)) 'mailbox_guid_invalid'
Require ($ImmutableId.Length -le 1024 -and $ImmutableId -notmatch '\s') 'opaque_id_invalid'
$s = $SnapshotJson | ConvertFrom-Json
Require ($s.schema -ceq 'mail-move-snapshot/v1') 'snapshot_schema'
Require ($null -ne $s.identity -and $null -ne $s.message -and $null -ne $s.folders) 'snapshot_fields'
Require ($s.identity.mailboxGuid -ieq $MailboxGuid) 'mailbox_mismatch'
Require ($s.identity.idType -ceq 'ImmutableId') 'immutable_id_unproven'
Require ($s.message.id -ceq $ImmutableId) 'message_mismatch'
Require (-not [string]::IsNullOrWhiteSpace($s.message.parentFolderId)) 'source_folder_missing'
Require (-not [string]::IsNullOrWhiteSpace($s.message.changeKey)) 'change_key_missing'
Require ($s.message.isDraft -ceq $false) 'draft_not_eligible'
Require ($s.message.flag.flagStatus -ceq 'notFlagged') 'open_flag_must_be_reviewed'
$targets = @($s.folders | Where-Object { $_.key -ceq $DestinationKey })
Require ($targets.Count -eq 1) 'destination_not_unique'
$target = $targets[0]
Require ($target.mailboxGuid -ieq $MailboxGuid) 'cross_mailbox_forbidden'
Require (-not [string]::IsNullOrWhiteSpace($target.id)) 'destination_folder_missing'
Require ($target.id -cne $s.message.parentFolderId) 'already_at_destination'

$plan = [ordered]@{
    mailboxGuid = $MailboxGuid
    immutableId = $ImmutableId
    fromFolderId = $s.message.parentFolderId
    toFolderId = $target.id
    destinationKey = $DestinationKey
    expectedChangeKey = $s.message.changeKey
}
$serialized = ConvertTo-Json -InputObject $plan -Depth 4 -Compress
$hasher = [Security.Cryptography.SHA256]::Create()
try {
    $digest = ([BitConverter]::ToString($hasher.ComputeHash([Text.Encoding]::UTF8.GetBytes($serialized)))).Replace('-', '').ToLowerInvariant()
} finally { $hasher.Dispose() }

[ordered]@{
    program = 'mail-move-plan/0.1.0-qa'
    status = 'PREVIEW_ONLY_UNATTESTED'
    fixture = [bool]$Fixture
    sourceTrust = 'CALLER_SUPPLIED_NOT_INDEPENDENTLY_ATTESTED'
    planSha256 = $digest
    preview = $plan
    predictedChanges = @('parentFolderId: fromFolderId -> toFolderId', 'new Graph REST id/changeKey are unknown until readback')
    requiredBeforeEffect = @('current mailbox/rights', 'fresh exact Graph GET with Prefer: IdType=ImmutableId', 'WWW case and folder policy', 'fresh source changeKey equality', 'complete chat preview and effect authority')
    requiredAfterEffect = @('original provider receipt', 'separate exact immutable message GET', 'folder and unchanged metadata validation', 'unknown-outcome reconciliation', 'WWW Gbrain per-source checkpoint')
    graphCalls = 0
    mutations = 0
    nativeMoveReceipt = $null
    independentReadback = $null
    executeSupported = $false
} | ConvertTo-Json -Depth 8 -Compress
