# Verkada Multiple Credentials Report

PowerShell script to identify Verkada Access users who have 2 or more access credentials assigned.

## Features

- Builds report for users with 2 or more credentials
- Saves progress to an incremental csv so the script can be resumed if stopped partway through
- Retry failed requests with exponential backoff
- Converts incremental CSV report to XLSX file when finished
- Logs any errors that are not solved with the retry logic
- Supports running on a custom subset of users, or all users in the org
- Displays real-time progress

## Example Output

| Name | Email | CompanyName | UserID | TotalCredentials | Link to User |
|------|-------|-------------|--------|------------------|--------------|
| User One | user1@example.com | Company A | 008c6e86-9a81-465d-8d4a-fd9b2a253e26 | 2 | https://yourorg.command.verkada.com/access/users/008c6e86-9a81-465d-8d4a-fd9b2a253e26/details |
| User Two | user2@example.com | Company B | 00c364c4-596b-42b9-b4a2-93d7ceac7335 | 2 | https://yourorg.command.verkada.com/access/users/00c364c4-596b-42b9-b4a2-93d7ceac7335/details |
| User Three | user3@example.com | Company C | 00fb7f0d-881d-426d-ab34-6b42196f5bfe | 3 | https://yourorg.command.verkada.com/access/users/00fb7f0d-881d-426d-ab34-6b42196f5bfe/details |

## Setup

1. Update `config.env` with the following values:
   - `ORG_SHORTNAME` - Your organization's subdomain (the part before `.command.verkada.com`)
   - `API_KEY` - Verkada API key with read access to Access Control endpoints
   - `MAX_USERS_TO_CHECK` - Set to `all` to check all users, or an integer to check a specific number per run

2. Install required PowerShell modules (the script will prompt to install these automatically if missing):
   - `verkadaModule` - [Verkada API client](https://github.com/bepsoccer/verkadaModule) for reading users
   - `ImportExcel` - For XLSX file generation

## Usage

### macOS

Install PowerShell (if not already installed), then run:
```bash
pwsh Generate-Report.ps1
```

Or double-click `macOS_run.command` in Finder.

### Windows / Linux / Any Platform with PowerShell

```powershell
pwsh ./Generate-Report.ps1
```

## How It Works

- **Progress Saving**: The script creates a log file tracking checked users. If stopped, re-run the script to resume from where it left off.
- **Retry Logic**: Failed API requests are retried with exponential backoff (starting at 2s, doubling each retry, max 32s). Requests that still fail after 10 attempts are logged to `access-users-error-log-[timestamp].csv`.
- **Output Files**:
  - Incremental CSV: `access-users-multiple-credentials-incremental-[timestamp].csv` (updated during run)
  - Final XLSX: `access-users-multiple-credentials-[timestamp].xlsx` (created when all users checked)
  - Log file: `access-users-log-[timestamp].csv` (tracks progress)
- When all users are checked, log and incremental CSV files are cleaned up automatically.
