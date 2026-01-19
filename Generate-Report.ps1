$PSStyle.OutputRendering = 'PlainText'

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$configFile = Join-Path $scriptDir "config.env"

if (-not (Test-Path $configFile)) {
    Write-Error "Configuration file not found: $configFile"
    exit 1
}

$configContent = Get-Content $configFile | Where-Object { $_ -notmatch '^\s*#' -and $_.Trim() -ne '' } -ErrorAction Stop
$config = ConvertFrom-StringData ($configContent -join "`n")

$MAX_USERS_TO_CHECK = $config.MAX_USERS_TO_CHECK
$ORG_SHORTNAME = $config.ORG_SHORTNAME
$API_KEY = $config.API_KEY

$timestamp = Get-Date -Format 'yyyy-MM-dd-HHmmss'
$LOG_FILE = Join-Path $scriptDir "access-users-log-$timestamp.csv"
$INCREMENTAL_CSV = Join-Path $scriptDir "access-users-multiple-credentials-incremental-$timestamp.csv"
$FINAL_XLSX = Join-Path $scriptDir "access-users-multiple-credentials-$timestamp.xlsx"
$ERROR_LOG = Join-Path $scriptDir "access-users-error-log-$timestamp.csv"

try {
    Import-Module verkadaModule -ErrorAction Stop
    Write-Host "verkadaModule imported successfully."
}
catch {
    Write-Warning "verkadaModule is required for this script, but it is not installed."
    $install = Read-Host "Would you like to install verkadaModule? (y/n)"
    if ($install -eq 'y' -or $install -eq 'Y') {
        Write-Host "Installing verkadaModule..."
        try {
            Install-Module -Name verkadaModule -Scope CurrentUser -Force
            Import-Module verkadaModule
            Write-Host "verkadaModule installed and imported successfully."
        }
        catch {
            Write-Error "Failed to install verkadaModule: $_"
            exit 1
        }
    }
    else {
        Write-Error "verkadaModule is required to run this script. Exiting."
        exit 1
    }
}

try {
    Import-Module ImportExcel -ErrorAction Stop
    Write-Host "ImportExcel module imported successfully."
}
catch {
    Write-Warning "ImportExcel module is required for this script, but it is not installed."
    $install = Read-Host "Would you like to install ImportExcel? (y/n)"
    if ($install -eq 'y' -or $install -eq 'Y') {
        Write-Host "Installing ImportExcel module..."
        try {
            Install-Module -Name ImportExcel -Scope CurrentUser -Force
            Import-Module ImportExcel
            Write-Host "ImportExcel module installed and imported successfully."
        }
        catch {
            Write-Error "Failed to install ImportExcel module: $_"
            exit 1
        }
    }
    else {
        Write-Error "ImportExcel module is required to run this script. Exiting."
        exit 1
    }
}

if (-not $API_KEY) {
    Write-Error "API_KEY is not set in config file"
    exit 1
}

try {
    Connect-Verkada -x_api_key $API_KEY
}
catch {
    Write-Error "Failed to connect to Verkada API: $_"
    exit 1
}

$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

$existingLogFiles = Get-ChildItem -Path $scriptDir -Filter "access-users-log-*.csv" | Sort-Object LastWriteTime -Descending
$checkedUserIds = @{}
$existingIncrementalCsv = $null

if ($existingLogFiles) {
    $LOG_FILE = $existingLogFiles[0].FullName
    Write-Host "Found existing log file: $LOG_FILE"
    Write-Host "Resuming from previous run..."

    # Read log file to get checked user IDs
    if (Test-Path $LOG_FILE) {
        $logEntries = Import-Csv -Path $LOG_FILE
        foreach ($entry in $logEntries) {
            $checkedUserIds[$entry.UserID] = $true
        }
        Write-Host "Already checked $($checkedUserIds.Count) users."
    }

    $logTimestamp = ($existingLogFiles[0].Name -replace 'access-users-log-', '' -replace '.csv', '')
    $potentialCsv = Join-Path $scriptDir "access-users-multiple-credentials-incremental-$logTimestamp.csv"
    if (Test-Path $potentialCsv) {
        $existingIncrementalCsv = $potentialCsv
        $INCREMENTAL_CSV = $potentialCsv
        Write-Host "Found existing incremental CSV: $INCREMENTAL_CSV"
    }

    $FINAL_XLSX = Join-Path $scriptDir "access-users-multiple-credentials-$logTimestamp.xlsx"
    $ERROR_LOG = Join-Path $scriptDir "access-users-error-log-$logTimestamp.csv"
} else {
    "Timestamp,UserID,UserName,Status" | Out-File -FilePath $LOG_FILE -Encoding utf8
    Write-Host "Created new log file: $LOG_FILE"

    "Timestamp,UserID,UserName" | Out-File -FilePath $ERROR_LOG -Encoding utf8
    Write-Host "Created error log file: $ERROR_LOG"
}

Write-Host "Retrieving Access Users list..."
try {
    $allUsers = Read-VerkadaAccessUsers -version v1 -x_verkada_auth_api $Global:verkadaConnection.x_verkada_auth_api
}
catch {
    Write-Error "Failed to retrieve access users: $_"
    exit 1
}

if (-not $allUsers) {
    Write-Warning "No access users found in the organization."
    exit
}

$uncheckedUsers = $allUsers | Where-Object { -not $checkedUserIds.ContainsKey($_.user_id) }

if ($MAX_USERS_TO_CHECK -eq 'all') {
    Write-Host "Checking credentials for remaining unchecked users..."
    $usersToProcess = $uncheckedUsers
} else {
    Write-Host "Checking credentials for up to $MAX_USERS_TO_CHECK unchecked users..."
    $usersToProcess = $uncheckedUsers | Select-Object -First $MAX_USERS_TO_CHECK
}

$usersToProcessCount = $usersToProcess.Count

if ($usersToProcessCount -eq 0) {
    Write-Host "All users have already been checked."
} else {
    Write-Host "Found $usersToProcessCount users to check (skipping $($checkedUserIds.Count) already checked)..."
}

# Iterate through users, get their full details, and append to CSV
# Use Get-VerkadaAccessUser to fetch details for each user found in the previous step.
$userCount = 0
$totalUsers = $usersToProcess.Count

$usersToProcess | ForEach-Object {
    $userCount++
    Write-Host "Checking user $userCount of $totalUsers..."
    $originalUser = $_
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

    $maxRetries = 10
    $retryCount = 0
    $initialDelay = 2
    $maxDelay = 32
    $userProcessed = $false

    while ($retryCount -lt $maxRetries -and -not $userProcessed) {
        try {
            $details = Get-VerkadaAccessUser -userId $_.user_id -ErrorAction Stop
            $cardCount = @($details.cards).Count
            $totalCredentials = $cardCount

            # Log the user check
            "$($timestamp),$($originalUser.user_id),$($originalUser.full_name),Checked" | Out-File -FilePath $LOG_FILE -Append -Encoding utf8

            if ($totalCredentials -ge 2) {
                Write-Host "  -> Found $($totalCredentials) credentials for $($originalUser.full_name)"

                # Append to incremental CSV
                $csvRow = [PSCustomObject]@{
                    Name            = $originalUser.full_name
                    Email           = $originalUser.email
                    CompanyName     = $originalUser.company_name
                    UserID          = $details.user_id
                    TotalCredentials = $totalCredentials
                    'Link to User'   = "https://$ORG_SHORTNAME.command.verkada.com/access/users/$($details.user_id)/details"
                }

                if (-not (Test-Path $INCREMENTAL_CSV)) {
                    $csvRow | Export-Csv -Path $INCREMENTAL_CSV -NoTypeInformation -Encoding utf8
                } else {
                    # Append without headers
                    $csvRow | ConvertTo-Csv -NoTypeInformation | Select-Object -Skip 1 | Out-File -FilePath $INCREMENTAL_CSV -Append -Encoding utf8
                }
            }

            $userProcessed = $true
        }
        catch {
            $retryCount++
            if ($retryCount -lt $maxRetries) {
                $waitTime = [math]::Min($initialDelay * [math]::Pow(2, $retryCount - 1), $maxDelay)
                Write-Host "  -> retrying $($originalUser.full_name)"
                Start-Sleep -Seconds $waitTime
            } else {
                "$($timestamp),$($originalUser.user_id),$($originalUser.full_name)" | Out-File -FilePath $ERROR_LOG -Append -Encoding utf8
            }
        }
    }
}

$totalOrgUsers = $allUsers.Count
$logEntries = Import-Csv -Path $LOG_FILE
$actualCheckedCount = $logEntries.Count

Write-Host "Progress: $actualCheckedCount users checked out of $totalOrgUsers total users in org."

if ($actualCheckedCount -ge $totalOrgUsers) {
    Write-Host "All users have been checked!"

    if (Test-Path $INCREMENTAL_CSV) {
        $csvData = Import-Csv -Path $INCREMENTAL_CSV
        $userCount = ($csvData | Measure-Object).Count

        if ($userCount -gt 0) {
            Write-Host "Found $userCount users with 2+ credentials. Converting to XLSX..."

            try {
                $csvData | Export-Excel -Path $FINAL_XLSX -AutoSize -AutoFilter
                Write-Host "Successfully converted to XLSX:"
                Write-Host "$FINAL_XLSX"

                Remove-Item -Path $LOG_FILE -Force
                Remove-Item -Path $INCREMENTAL_CSV -Force
                Write-Host "Cleaned up log and temporary CSV files."
            }
            catch {
                Write-Warning "Failed to convert CSV to XLSX: $_"
                Write-Host "Incremental CSV is available at: $INCREMENTAL_CSV"
            }
        } else {
            Write-Host "No users found with 2 or more access credentials."

            Remove-Item -Path $LOG_FILE -Force
            if (Test-Path $INCREMENTAL_CSV) {
                Remove-Item -Path $INCREMENTAL_CSV -Force
            }
            Write-Host "Cleaned up log files."
        }
    } else {
        Write-Host "No users found with 2 or more access credentials."

        Remove-Item -Path $LOG_FILE -Force
        Write-Host "Cleaned up log file."
    }
} else {
    Write-Host "Run the script again to continue checking the remaining users."
    Write-Host "Log file: $LOG_FILE"
}

$stopwatch.Stop()
$elapsed = $stopwatch.Elapsed

if ($elapsed.TotalHours -ge 1) {
    $timeString = "{0}h {1}m {2}s" -f [int]$elapsed.TotalHours, $elapsed.Minutes, $elapsed.Seconds
} elseif ($elapsed.TotalMinutes -ge 1) {
    $timeString = "{0}m {1}s" -f [int]$elapsed.TotalMinutes, $elapsed.Seconds
} else {
    $timeString = "{0}s" -f [int]$elapsed.TotalSeconds
}

if (Test-Path $ERROR_LOG) {
    $errorLines = Get-Content -Path $ERROR_LOG
    if ($errorLines.Count -le 1) {
        Remove-Item -Path $ERROR_LOG -Force
        Write-Host "No errors logged, removed error log file."
    }
}

Write-Host "`nTotal time: $timeString"
