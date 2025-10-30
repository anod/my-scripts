# ms_todo_to_todoist.ps1
# Export Microsoft To Do tasks to Todoist-compatible CSV (Type,Content,Priority,Due Date,Due Time,Description)
#
# Requirements:
#   - PowerShell 7+ recommended
#   - Environment variables:
#         MS_TODO_CLIENT_ID   (Azure AD app registration with Tasks.Read or Tasks.ReadWrite delegated permission)
#         MS_TODO_TENANT_ID   (tenant id or 'common', optional, defaults to 'common')
#         MS_TODO_SCOPE_READWRITE=1 (optional; if set uses Tasks.ReadWrite scope, else Tasks.Read)
#   - No external modules required (pure REST + device code flow)
#
# Usage examples:
#   pwsh ./ms_todo_to_todoist.ps1 -Output tasks.csv
#   pwsh ./ms_todo_to_todoist.ps1 -Output tasks.csv -IncludeCompleted -IncludeCheckedChecklistItems -ListPrefix -PriorityMap "high=4,normal=2,low=1"
#   pwsh ./ms_todo_to_todoist.ps1 -Output tasks.csv -NoChecklists
#
# Notes:
#   - By default completed tasks are excluded (use -IncludeCompleted to include)
#   - Checklist items become separate rows unless -NoChecklists is specified
#   - Completed checklist items excluded unless -IncludeCheckedChecklistItems is specified
#   - Priority mapping defaults: high=4, normal=1, low=1 (Todoist: 1 lowest .. 4 highest)
#   - List name appended in Description; optionally prefixed in Content with -ListPrefix
#   - Subtasks (checklist items) appear as "Parent Title > Checklist Item" in Content
#
# Limitations:
#   - Attachments are not exported (Graph only notes hasAttachments flag)
#   - Pagination handled for large lists/tasks
#
# Exit codes:
#   0 success
#   1 missing env vars / auth failure
#   2 invalid parameter values
#   3 Graph API error
#
# CSV Columns:
#   Type,Content,Priority,Due Date,Due Time,Description

[CmdletBinding()] param(
    [Parameter(Mandatory=$true)] [string] $Output,
    [switch] $IncludeCompleted,
    [switch] $NoChecklists,
    [switch] $IncludeCheckedChecklistItems,
    [switch] $ListPrefix,
    [string] $PriorityMap,
    [switch] $VerboseLogging
)

function Write-Log {
    param([string] $Message)
    if ($VerboseLogging) { Write-Host "[info] $Message" }
}

# Parse priority map string (e.g. "high=4,normal=1,low=1")
function Parse-PriorityMap {
    param([string] $MapString)
    $default = @{ high = 4; normal = 1; low = 1 }
    if (-not $MapString) { return $default }
    $pairs = $MapString -split ',' | Where-Object { $_.Trim() -ne '' }
    foreach ($p in $pairs) {
        if ($p -notmatch '=') { Write-Error "Invalid priority map segment '$p'"; exit 2 }
        $k,$v = $p.Split('=',2)
        $k = $k.Trim().ToLower()
        $v = $v.Trim()
        if ($v -notmatch '^[1-4]$') { Write-Error "Priority must be 1..4: '$v'"; exit 2 }
        $default[$k] = [int]$v
    }
    return $default
}

$PriorityMapping = Parse-PriorityMap -MapString $PriorityMap

# Auth: Device Code Flow
$clientId = $env:MS_TODO_CLIENT_ID
if (-not $clientId) { Write-Error "MS_TODO_CLIENT_ID not set"; exit 1 }
$tenant = $env:MS_TODO_TENANT_ID
if (-not $tenant) { $tenant = 'common' }
$scope = ($env:MS_TODO_SCOPE_READWRITE) ? 'https://graph.microsoft.com/Tasks.ReadWrite' : 'https://graph.microsoft.com/Tasks.Read'

$deviceCodeEndpoint = "https://login.microsoftonline.com/$tenant/oauth2/v2.0/devicecode"
$tokenEndpoint      = "https://login.microsoftonline.com/$tenant/oauth2/v2.0/token"

Write-Log "Requesting device code..."
try {
    $dcResponse = Invoke-RestMethod -Method Post -Uri $deviceCodeEndpoint -Body @{ client_id=$clientId; scope=$scope } -ErrorAction Stop
} catch {
    Write-Error "Failed to initiate device flow: $_"; exit 1
}

Write-Host "To authenticate, visit: $($dcResponse.verification_uri)" -ForegroundColor Cyan
Write-Host "Enter code: $($dcResponse.user_code)" -ForegroundColor Yellow

$deviceCode = $dcResponse.device_code
$interval   = [int]$dcResponse.interval
$expiresIn  = [int]$dcResponse.expires_in
$startTime  = Get-Date
$token = $null

Write-Log "Polling for token..."
while (-not $token) {
    if ((Get-Date) -gt $startTime.AddSeconds($expiresIn)) { Write-Error "Device code expired before authentication"; exit 1 }
    Start-Sleep -Seconds $interval
    try {
        $tokenResponse = Invoke-RestMethod -Method Post -Uri $tokenEndpoint -Body @{ grant_type='urn:ietf:params:oauth:grant-type:device_code'; client_id=$clientId; device_code=$deviceCode } -ErrorAction Stop
        if ($tokenResponse.access_token) { $token = $tokenResponse.access_token; break }
    } catch {
        # Expecting authorization_pending until user finishes
        $errJson = $_.ErrorDetails.Message | ConvertFrom-Json -ErrorAction SilentlyContinue
        $error = $errJson.error
        if ($error -eq 'authorization_pending') { continue }
        if ($error -eq 'slow_down') { $interval += 5; continue }
        Write-Error "Token polling error: $error"; exit 1
    }
}
Write-Log "Token acquired"

$authHeader = @{ Authorization = "Bearer $token" }

function Invoke-GraphPaged {
    param([string] $Url)
    $results = @()
    while ($Url) {
        Write-Log "GET $Url"
        try {
            $resp = Invoke-RestMethod -Uri $Url -Headers $authHeader -ErrorAction Stop
        } catch {
            Write-Error "Graph request failed: $_"; exit 3
        }
        if ($resp.value) { $results += $resp.value }
        $Url = $resp.'@odata.nextLink'
    }
    return $results
}

Write-Log "Fetching lists..."
$lists = Invoke-GraphPaged -Url 'https://graph.microsoft.com/v1.0/me/todo/lists'
Write-Log "Found $($lists.Count) lists"

$allRows = @()

function Parse-Due {
    param([object] $Due)
    if (-not $Due -or -not $Due.dateTime) { return '', '' }
    $raw = $Due.dateTime
    try {
        $dtObj = [DateTime]::Parse($raw)
    } catch { return '', '' }
    $dateStr = $dtObj.ToString('yyyy-MM-dd')
    $timeStr = if ($dtObj.TimeOfDay.TotalSeconds -gt 0) { $dtObj.ToString('HH:mm') } else { '' }
    return $dateStr, $timeStr
}

function Sanitize([string] $Text) {
    if (-not $Text) { return '' }
    return ($Text -replace '\r', ' ' -replace '\n', ' ').Trim()
}

function Importance-ToPriority {
    param([string] $Importance)
    if (-not $Importance) { $Importance = 'normal' }
    $Importance = $Importance.ToLower()
    return $PriorityMapping[$Importance]
}

foreach ($list in $lists) {
    $listId = $list.id
    $listName = if ($list.displayName) { $list.displayName } else { 'Unnamed List' }
    Write-Log "Fetching tasks for list '$listName'"
    $tasks = Invoke-GraphPaged -Url "https://graph.microsoft.com/v1.0/me/todo/lists/$listId/tasks?$expand=checklistItems"
    foreach ($task in $tasks) {
        $status = $task.status
        if ($status -eq 'completed' -and -not $IncludeCompleted) { continue }
        $title = Sanitize $task.title
        $importance = $task.importance
        $priority = Importance-ToPriority $importance
        $dueDate,$dueTime = Parse-Due $task.dueDateTime
        $bodyContent = if ($task.body.content) { Sanitize $task.body.content } else { '' }
        $descParts = @()
        if ($bodyContent) { $descParts += $bodyContent }
        if ($task.hasAttachments) { $descParts += '[Has Attachments]' }
        $descParts += "List: $listName"
        $description = $descParts -join "`n"
        $content = if ($ListPrefix) { "$listName | $title" } else { $title }
        $allRows += [PSCustomObject]@{
            Type        = 'task'
            Content     = $content
            Priority    = $priority
            'Due Date'  = $dueDate
            'Due Time'  = $dueTime
            Description = $description
        }
        if (-not $NoChecklists -and $task.checklistItems) {
            foreach ($cl in $task.checklistItems) {
                $isChecked = $cl.isChecked
                if ($isChecked -and -not $IncludeCheckedChecklistItems) { continue }
                $clTitle = Sanitize $cl.displayName
                $allRows += [PSCustomObject]@{
                    Type        = 'task'
                    Content     = "$title > $clTitle"
                    Priority    = $priority
                    'Due Date'  = $dueDate
                    'Due Time'  = $dueTime
                    Description = "Subtask from '$title' | List: $listName"
                }
            }
        }
    }
}

Write-Log "Writing $($allRows.Count) rows to $Output"
try {
    $allRows | Export-Csv -Path $Output -NoTypeInformation -Encoding UTF8
} catch {
    Write-Error "Failed to write CSV: $_"; exit 3
}

Write-Host "Exported $($allRows.Count) rows to $Output" -ForegroundColor Green
exit 0
