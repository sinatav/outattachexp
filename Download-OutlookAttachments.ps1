<#
.SYNOPSIS
  Interactive Outlook attachment downloader for Windows.

.DESCRIPTION
  Uses the classic Outlook desktop client's COM interface, so it inherits
  the signed-in session — no Graph, no app registration, no IMAP, no
  passwords. Prompts for folder, date range, sender filter, file
  extensions, and a destination directory, then saves matching
  attachments.

.PARAMETER (none)
  Interactive only.

.NOTES
  Requirements:
    - Classic Outlook for Windows (NOT "New Outlook"). Must be installed
      and signed in with the target account.
    - Windows PowerShell 5.1 (built-in) or PowerShell 7.

  How to run:
    1. Save this file to your machine, e.g. Desktop.
    2. Right-click -> "Run with PowerShell"
       OR open PowerShell and run:
         powershell -ExecutionPolicy Bypass -File "$HOME\Desktop\Download-OutlookAttachments.ps1"
    3. Answer the prompts.
#>

[CmdletBinding()]
param()

function Read-Default {
    param([string]$Label, [string]$Default = "")
    $suffix = if ($Default) { " [$Default]" } else { "" }
    $value = Read-Host "$Label$suffix"
    if ([string]::IsNullOrWhiteSpace($value)) { return $Default }
    return $value.Trim()
}

function ConvertTo-SafeFileName {
    param([string]$Name)
    $invalid = [IO.Path]::GetInvalidFileNameChars()
    foreach ($c in $invalid) { $Name = $Name.Replace($c, '_') }
    $Name = $Name.Trim().TrimEnd('.')
    if ([string]::IsNullOrWhiteSpace($Name)) { $Name = 'attachment' }
    return $Name
}

function Get-UniquePath {
    param([string]$Dir, [string]$Name)
    $candidate = Join-Path $Dir $Name
    if (-not (Test-Path -LiteralPath $candidate)) { return $candidate }
    $stem = [IO.Path]::GetFileNameWithoutExtension($Name)
    $ext  = [IO.Path]::GetExtension($Name)
    $i = 1
    while ($true) {
        $candidate = Join-Path $Dir ("{0} ({1}){2}" -f $stem, $i, $ext)
        if (-not (Test-Path -LiteralPath $candidate)) { return $candidate }
        $i++
    }
}

function Resolve-OutlookFolder {
    param($Namespace, [string]$Path)
    $inbox = $Namespace.GetDefaultFolder(6) # olFolderInbox
    if ([string]::IsNullOrWhiteSpace($Path) -or $Path -ieq 'Inbox') { return $inbox }
    $parts = $Path -split '[/\\]'
    if ($parts[0] -ieq 'Inbox') {
        $current = $inbox
        $start = 1
    } else {
        $current = $inbox.Parent.Folders.Item($parts[0])
        $start = 1
    }
    for ($i = $start; $i -lt $parts.Length; $i++) {
        $current = $current.Folders.Item($parts[$i])
    }
    return $current
}

Write-Host "Outlook attachment downloader (Windows / COM)"
Write-Host ("-" * 50)

$folderPath = Read-Default "Folder (e.g. Inbox, Inbox/SubFolder)" "Inbox"
$fromStr    = Read-Default "From date (YYYY-MM-DD, blank = no lower bound)"
$toStr      = Read-Default "To date   (YYYY-MM-DD, blank = no upper bound)"
$sender     = Read-Default "Sender filter (email = exact, plain text = substring; blank = any)"
$extsStr    = Read-Default "File extensions (comma-separated, e.g. pdf,xlsx; blank = all)"
$dest       = Read-Default "Download destination" (Join-Path $env:USERPROFILE "Downloads\outlook-attachments")

$from = if ($fromStr) { [datetime]::ParseExact($fromStr, "yyyy-MM-dd", $null) } else { $null }
$to   = if ($toStr)   { [datetime]::ParseExact($toStr,   "yyyy-MM-dd", $null) } else { $null }
$exts = @(
    $extsStr.Split(',') |
        ForEach-Object { $_.Trim().ToLower().TrimStart('.') } |
        Where-Object { $_ }
)

if (-not (Test-Path -LiteralPath $dest)) {
    New-Item -ItemType Directory -Path $dest -Force | Out-Null
}

Write-Host ""
Write-Host "Connecting to Outlook..."
try {
    $outlook   = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
} catch {
    Write-Error "Could not connect to Outlook. Is the classic Outlook desktop client installed and running? Details: $_"
    exit 1
}

try {
    $folder = Resolve-OutlookFolder -Namespace $namespace -Path $folderPath
} catch {
    Write-Error "Folder '$folderPath' not found. Details: $_"
    exit 2
}

$items = $folder.Items
$items.Sort("[ReceivedTime]", $true)

# Narrow with Restrict for the date range — much faster than enumerating
# every item, especially for large mailboxes.
$clauses = @()
if ($from) { $clauses += "[ReceivedTime] >= '" + $from.ToString('g') + "'" }
if ($to)   { $clauses += "[ReceivedTime] <= '" + $to.AddDays(1).AddSeconds(-1).ToString('g') + "'" }
if ($clauses.Count -gt 0) {
    $items = $items.Restrict([string]::Join(' AND ', $clauses))
}

$scanned = 0
$saved   = 0

foreach ($mail in $items) {
    $scanned++
    if ($mail.Class -ne 43) { continue } # 43 = olMail; skip meetings/tasks/etc.
    if (-not $mail.Attachments -or $mail.Attachments.Count -eq 0) { continue }

    if ($sender) {
        $addr = ""; $name = ""
        try { $addr = $mail.SenderEmailAddress } catch {}
        try { $name = $mail.SenderName } catch {}
        $isEmail = $sender.Contains("@")
        $matches = if ($isEmail) {
            $addr -ieq $sender
        } else {
            ($addr -like "*$sender*") -or ($name -like "*$sender*")
        }
        if (-not $matches) { continue }
    }

    foreach ($att in $mail.Attachments) {
        # Type 1 = olByValue (a real file). Skip OLE/embedded items.
        if ($att.Type -ne 1) { continue }
        $fname = $att.FileName
        if (-not $fname) { continue }

        if ($exts.Count -gt 0) {
            $ext = [IO.Path]::GetExtension($fname).TrimStart('.').ToLower()
            if (-not ($exts -contains $ext)) { continue }
        }

        $safe = ConvertTo-SafeFileName $fname
        $outPath = Get-UniquePath -Dir $dest -Name $safe
        try {
            $att.SaveAsFile($outPath)
            $saved++
            $subject = if ($mail.Subject) { $mail.Subject } else { "(no subject)" }
            if ($subject.Length -gt 80) { $subject = $subject.Substring(0, 80) }
            Write-Host ("  [{0}] {1} | {2}" -f $scanned, $mail.SenderEmailAddress, $subject)
            Write-Host ("        -> {0}" -f $outPath)
        } catch {
            Write-Warning ("Failed to save '{0}': {1}" -f $fname, $_)
        }
    }
}

Write-Host ""
Write-Host ("Done. Scanned {0} message(s), saved {1} file(s) to {2}" -f $scanned, $saved, $dest)
Write-Host ""
Read-Host "Press Enter to close"
