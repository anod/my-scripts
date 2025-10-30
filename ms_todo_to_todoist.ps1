# Export Microsoft To Do tasks to a Todoist template-compatible CSV.
# Template columns:
# TYPE,CONTENT,DESCRIPTION,PRIORITY,INDENT,AUTHOR,RESPONSIBLE,DATE,DATE_LANG,TIMEZONE,DURATION,DURATION_UNIT,DEADLINE,DEADLINE_LANG
# AUTHOR populated from your GitHub profile (Alex Gavrishev) if available.
# INDENT: 1 for top-level tasks, 2 for checklist-derived subtasks.
# DATE: human-friendly due (YYYY-MM-DD HH:MM if time exists, else YYYY-MM-DD). DEADLINE: due date (YYYY-MM-DD).
# PRIORITY mapping default: high=4 normal=1 low=1 (Todoist: 1 lowest .. 4 highest).
# Checklist items become separate tasks with TYPE=task and INDENT=2 ("Parent > Item").
# Completed tasks excluded unless -IncludeCompleted specified.
# Completed checklist items excluded unless -IncludeCheckedChecklistItems specified.

[CmdletBinding()] param(
    [Parameter(Mandatory=$true)] [string] $Output,
    [switch] $IncludeCompleted,
    [switch] $IncludeCheckedChecklistItems,
    [switch] $NoChecklists,
    [string] $PriorityMap,
    [switch] $VerboseLogging,
    [switch] $SplitLists,            # When set, create one CSV per list inside the -Output directory (from live Graph fetch)
    [string] $SplitExisting          # Path to an existing combined CSV to split per list (no Graph calls)
)

function Write-Log { param([string] $Message) if ($VerboseLogging) { Write-Host "[info] $Message" } }

function Get-SafeFileName {
    param([string] $Name)
    if (-not $Name -or $Name.Trim() -eq '') { $Name = 'Unnamed List' }
    $invalid = [IO.Path]::GetInvalidFileNameChars() -join ''
    $regex = "[" + [Regex]::Escape($invalid) + "]"
    $safe = ($Name -replace $regex,'_').Trim()
    if ($safe.Length -gt 100) { $safe = $safe.Substring(0,100) }
    if (-not $safe) { $safe = 'List' }
    return $safe
}

function Resolve-PerListOutputPath {
    param(
        [string] $BaseDirectory,
        [string] $ListName,
        [hashtable] $Used
    )
    $base = Get-SafeFileName -Name $ListName
    $candidate = $base
    $i = 1
    while ($Used.ContainsKey($candidate)) {
        $i++
        $candidate = "$base-$i"
    }
    $Used[$candidate] = $true
    return Join-Path -Path $BaseDirectory -ChildPath ("$candidate.csv")
}

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

# Author info (from previous profile fetch)
$AuthorName = 'Alex Gavrishev'
if (-not $AuthorName) { $AuthorName = 'anod' }

# If splitting an existing CSV, perform that action and exit early.
if ($SplitExisting) {
    if ($SplitLists) { Write-Error "-SplitLists cannot be combined with -SplitExisting"; exit 5 }
    if (-not (Test-Path -LiteralPath $SplitExisting)) { Write-Error "Split source CSV not found: $SplitExisting"; exit 5 }
    # Output must be a directory
    if (-not (Test-Path -LiteralPath $Output)) {
        Write-Log "Creating output directory '$Output'"
        try { New-Item -ItemType Directory -Path $Output -Force | Out-Null } catch { Write-Error "Cannot create output directory '$Output': $_"; exit 5 }
    } elseif (-not (Get-Item -LiteralPath $Output).PSIsContainer) {
        Write-Error "When using -SplitExisting, -Output must be a directory"; exit 5 }

    Write-Log "Loading existing CSV: $SplitExisting"
    try { $existingRows = Import-Csv -Path $SplitExisting -ErrorAction Stop } catch { Write-Error "Failed to read CSV '$SplitExisting': $_"; exit 5 }
    if (-not $existingRows -or $existingRows.Count -eq 0) { Write-Error "No rows in source CSV"; exit 5 }

    $usedNames = @{}
    $groups = @{}
    foreach ($r in $existingRows) {
        $desc = $r.DESCRIPTION
        $listLine = ($desc -split "`n" | Where-Object { $_ -match '^List:\s*' } | Select-Object -Last 1)
        $listName = if ($listLine) { ($listLine -replace '^List:\s*','').Trim() } else { 'Unknown List' }
        if (-not $groups.ContainsKey($listName)) { $groups[$listName] = @() }
        $groups[$listName] += $r
    }
    Write-Log "Identified $($groups.Keys.Count) lists from existing CSV"
    foreach ($ln in $groups.Keys) {
        $outPath = Resolve-PerListOutputPath -BaseDirectory $Output -ListName $ln -Used $usedNames
        $listRows = $groups[$ln]
        Write-Log "Writing list '$ln' rows: $($listRows.Count) -> $outPath"
        try { $listRows | Export-Csv -Path $outPath -NoTypeInformation -Encoding UTF8 } catch { Write-Error "Failed writing '$ln': $_"; exit 5 }
    }
    Write-Host "Split existing CSV into $($groups.Keys.Count) files under $Output" -ForegroundColor Green
    exit 0
}

# Auth (Device Code Flow) only needed for live fetch
$clientId = $env:MS_TODO_CLIENT_ID
if (-not $clientId) { Write-Error "MS_TODO_CLIENT_ID not set"; exit 1 }
$tenant  = 'common'
$scope   = ($env:MS_TODO_SCOPE_READWRITE) ? 'https://graph.microsoft.com/Tasks.ReadWrite' : 'https://graph.microsoft.com/Tasks.Read'

$deviceCodeEndpoint = "https://login.microsoftonline.com/$tenant/oauth2/v2.0/devicecode"
$tokenEndpoint      = "https://login.microsoftonline.com/$tenant/oauth2/v2.0/token"

Write-Log "Initiating device code flow"
try {
    $dcResponse = Invoke-RestMethod -Method Post -Uri $deviceCodeEndpoint -Body @{ client_id=$clientId; scope=$scope } -ErrorAction Stop
} catch { Write-Error "Device code request failed: $_"; exit 1 }

Write-Host "Authenticate at: $($dcResponse.verification_uri)" -ForegroundColor Cyan
Write-Host "Enter code: $($dcResponse.user_code)" -ForegroundColor Yellow

$deviceCode = $dcResponse.device_code
$interval   = [int]$dcResponse.interval
$expiresIn  = [int]$dcResponse.expires_in
$startTime  = Get-Date
$token      = $null

while (-not $token) {
    if ((Get-Date) -gt $startTime.AddSeconds($expiresIn)) { Write-Error "Device code expired"; exit 1 }
    Start-Sleep -Seconds $interval
    try {
        $tokenResponse = Invoke-RestMethod -Method Post -Uri $tokenEndpoint -Body @{
            grant_type  = 'urn:ietf:params:oauth:grant-type:device_code'
            client_id   = $clientId
            device_code = $deviceCode
        } -ErrorAction Stop
        if ($tokenResponse.access_token) { $token = $tokenResponse.access_token }
    } catch {
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
        try { $resp = Invoke-RestMethod -Uri $Url -Headers $authHeader -ErrorAction Stop }
        catch { Write-Error "Graph request failed: $_"; exit 3 }
        if ($resp.value) { $results += $resp.value }
        $Url = $resp.'@odata.nextLink'
    }
    return $results
}

function Sanitize([string] $Text) { if (-not $Text) { return '' }; ($Text -replace '\r',' ' -replace '\n',' ').Trim() }

function Importance-ToPriority {
    param([string] $Importance)
    if (-not $Importance) { $Importance = 'normal' }
    $PriorityMapping[$Importance.ToLower()]
}

function Parse-Due {
    param([object] $Due)
    if (-not $Due -or -not $Due.dateTime) { return '', '', '' }
    $raw = $Due.dateTime
    try { $dtObj = [DateTime]::Parse($raw) } catch { return '', '', '' }
    $date = $dtObj.ToString('yyyy-MM-dd')
    $time = if ($dtObj.TimeOfDay.TotalSeconds -gt 0) { $dtObj.ToString('HH:mm') } else { '' }
    $display = if ($time) { "$date $time" } else { $date }
    return $date, $time, $display
}

Write-Log "Fetching lists"
$lists = Invoke-GraphPaged -Url 'https://graph.microsoft.com/v1.0/me/todo/lists'
Write-Log "Lists: $($lists.Count)"

$rows = @()

# Prepare directory if splitting
if ($SplitLists) {
    # Treat -Output as directory path
    if (-not (Test-Path -LiteralPath $Output)) {
        Write-Log "Creating output directory '$Output'"
        try { New-Item -ItemType Directory -Path $Output -Force | Out-Null } catch { Write-Error "Cannot create output directory '$Output': $_"; exit 4 }
    } elseif (-not (Get-Item -LiteralPath $Output).PSIsContainer) {
        Write-Error "When using -SplitLists, -Output must point to a directory"; exit 4
    }
}

$usedNames = @{}

foreach ($list in $lists) {
    $listId = $list.id
    $listName = if ($list.displayName) { $list.displayName } else { 'Unnamed List' }
    Write-Log "Processing list: $listName ($listId)"
    $tasks = Invoke-GraphPaged -Url "https://graph.microsoft.com/v1.0/me/todo/lists/$listId/tasks?$expand=checklistItems"
    $listRows = @()
    foreach ($task in $tasks) {
        if ($task.status -eq 'completed' -and -not $IncludeCompleted) { continue }
        $title = Sanitize $task.title
        $priority = Importance-ToPriority $task.importance
        $dueDate,$dueTime,$dateDisplay = Parse-Due $task.dueDateTime

        $bodyContent = if ($task.body.content) { Sanitize $task.body.content } else { '' }
        $descParts = @()
        if ($bodyContent) { $descParts += $bodyContent }
        if ($task.hasAttachments) { $descParts += '[Has Attachments]' }
        $descParts += "List: $listName"
        $description = $descParts -join "`n"

        $taskObj = [PSCustomObject]@{
            TYPE          = 'task'
            CONTENT       = $title
            DESCRIPTION   = $description
            PRIORITY      = $priority
            INDENT        = 1
            AUTHOR        = $AuthorName
            RESPONSIBLE   = ''
            DATE          = $dateDisplay
            DATE_LANG     = 'en'
            TIMEZONE      = ''
            DURATION      = ''
            DURATION_UNIT = ''
            DEADLINE      = $dueDate
            DEADLINE_LANG = if ($dueDate) { 'en' } else { '' }
        }

        if ($SplitLists) { $listRows += $taskObj } else { $rows += $taskObj }

        if (-not $NoChecklists -and $task.checklistItems) {
            foreach ($cl in $task.checklistItems) {
                if ($cl.isChecked -and -not $IncludeCheckedChecklistItems) { continue }
                $clTitle = Sanitize $cl.displayName
                $clObj = [PSCustomObject]@{
                    TYPE          = 'task'
                    CONTENT       = "$title > $clTitle"
                    DESCRIPTION   = "Subtask from '$title' | List: $listName"
                    PRIORITY      = $priority
                    INDENT        = 2
                    AUTHOR        = $AuthorName
                    RESPONSIBLE   = ''
                    DATE          = $dateDisplay
                    DATE_LANG     = 'en'
                    TIMEZONE      = ''
                    DURATION      = ''
                    DURATION_UNIT = ''
                    DEADLINE      = $dueDate
                    DEADLINE_LANG = if ($dueDate) { 'en' } else { '' }
                }
                if ($SplitLists) { $listRows += $clObj } else { $rows += $clObj }
            }
        }
    }

    if ($SplitLists) {
        $outPath = Resolve-PerListOutputPath -BaseDirectory $Output -ListName $listName -Used $usedNames
        Write-Log "Writing list '$listName' rows: $($listRows.Count) -> $outPath"
        try { $listRows | Export-Csv -Path $outPath -NoTypeInformation -Encoding UTF8 } catch { Write-Error "Failed writing list '$listName' to '$outPath': $_"; exit 3 }
    }
}

if ($SplitLists) {
    Write-Host "Exported per-list CSV files to $Output" -ForegroundColor Green
} else {
    Write-Log "Writing $($rows.Count) rows"
    try {
        $rows | Export-Csv -Path $Output -NoTypeInformation -Encoding UTF8
    } catch {
        Write-Error "Failed writing CSV: $_"; exit 3
    }
    Write-Host "Exported $($rows.Count) rows to $Output" -ForegroundColor Green
}
exit 0