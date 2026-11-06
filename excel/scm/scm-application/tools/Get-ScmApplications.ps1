# Get-ScmApplications.ps1
# Download all applications (predefined App-IDs + custom) from SCM to CSV.
# Every page (1000 items) is appended to the CSV as soon as it is fetched,
# so a stopped run can be resumed with -StartOffset.
# Usage (PowerShell):
#   powershell -ExecutionPolicy Bypass -File .\Get-ScmApplications.ps1 -ClientId "xxx@yyy.iam.panserviceaccount.com" -TsgId "1234567890"
#   resume:  ... -StartOffset 53000      (the next offset shown in the last progress line)
#   by snippet instead of folder:  ... -Snippet "predefined-snippet"
#   (client_secret is asked interactively; -Folder defaults to Shared; -Snippet overrides -Folder)
param(
    [Parameter(Mandatory = $true)][string]$ClientId,
    [Parameter(Mandatory = $true)][string]$TsgId,
    [string]$Folder = "Shared",
    [string]$Snippet = "",
    [string]$Path = "/sse/config/v1/applications",
    [string]$OutFile = ".\applications.csv",
    [int]$StartOffset = 0,
    [int]$Limit = 1000,
    [int]$Retry = 3,
    [string]$BaseUrl = "https://api.sase.paloaltonetworks.com",
    [string]$AuthUrl = "https://auth.apps.paloaltonetworks.com/oauth2/access_token",
    [string]$Proxy = ""
)
$ErrorActionPreference = "Stop"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$sec = Read-Host "client_secret" -AsSecureString
$secret = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($sec))
$px = @{}
if ($Proxy -ne "") { $px = @{ Proxy = $Proxy } }

function Get-Token {
    $basic = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("$($ClientId):$($secret)"))
    $t = Invoke-RestMethod -Method Post -Uri $AuthUrl @px -Headers @{ Authorization = "Basic $basic" } `
        -ContentType "application/x-www-form-urlencoded" -Body "grant_type=client_credentials&scope=tsg_id:$TsgId"
    $script:tokenAt = Get-Date
    return $t.access_token
}

# Fetch one page; retry on errors (new token each retry)
function Get-Page([int]$off) {
    if ($Snippet -ne "") { $q = "snippet=$([uri]::EscapeDataString($Snippet))" } else { $q = "folder=$([uri]::EscapeDataString($Folder))" }
    $u = "$BaseUrl$($Path)?$q&limit=$Limit&offset=$off"
    for ($i = 1; $i -le $Retry; $i++) {
        try {
            return Invoke-RestMethod -Uri $u @px -Headers @{ Authorization = "Bearer $script:token"; Accept = "application/json" }
        } catch {
            Write-Warning ("offset {0}: attempt {1}/{2} failed: {3}" -f $off, $i, $Retry, $_.Exception.Message)
            if ($i -eq $Retry) { throw }
            Start-Sleep -Seconds (5 * $i)
            $script:token = Get-Token
        }
    }
}

if ($StartOffset -eq 0 -and (Test-Path $OutFile)) { Remove-Item $OutFile }
$script:token = Get-Token
$offset = $StartOffset
$total = 0
do {
    # refresh the token every 10 minutes (it expires after about 15)
    if (((Get-Date) - $tokenAt).TotalMinutes -gt 10) { $script:token = Get-Token }
    $r = Get-Page $offset
    $page = foreach ($a in $r.data) {
        [pscustomobject]@{
            name        = $a.name
            id          = $a.id
            folder      = $a.folder
            snippet     = $a.snippet
            category    = $a.category
            subcategory = $a.subcategory
            technology  = $a.technology
            risk        = $a.risk
        }
    }
    $n = @($r.data).Count
    if ($n -gt 0) { $page | Export-Csv -Path $OutFile -NoTypeInformation -Encoding UTF8 -Append }
    $total += $n
    $offset += $Limit
    Write-Host ("saved {0} rows this run / total {1} / next offset {2}" -f $total, $r.total, $offset)
} while ($n -eq $Limit)

Write-Host "Done. $total rows saved to $OutFile"
