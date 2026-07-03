<#
.SYNOPSIS
    Downloads AL symbols from global NuGet sources.
    Replicates the VS Code "AL: Download Symbols from global sources" command.

.DESCRIPTION
    Reads app.json, resolves all dependencies (including transitive), and downloads
    symbol .app packages from the Microsoft BC AppSource NuGet feed to .alpackages.

.PARAMETER WorkspacePath
    Root of the AL workspace. Defaults to the script's own directory.

.PARAMETER Country
    Symbols country/region code. Default: 'w1' (worldwide).

.PARAMETER OutputPath
    Destination folder for downloaded .app files.
    Defaults to <WorkspacePath>\.alpackages.

.PARAMETER Force
    Re-download packages even if they already exist in the output folder.

.EXAMPLE
    .\Download-AlSymbols.ps1

.EXAMPLE
    .\Download-AlSymbols.ps1 -Country w1 -Force
#>

param(
    [string] $WorkspacePath = $PSScriptRoot,
    [string] $Country       = "w1",
    [string] $OutputPath    = "",
    [switch] $Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ─────────────────────────────────────────────────────────────────────────────
# NuGet feed (AppSource symbols — same feed used by the AL extension)
# Endpoints are resolved dynamically from the service index because Azure
# DevOps Artifacts uses internal GUIDs that differ from the named URL.
# ─────────────────────────────────────────────────────────────────────────────
$FeedIndexUrl = "https://pkgs.dev.azure.com/dynamicssmb2/DynamicsBCPublicFeeds/_packaging/AppSourceSymbols/nuget/v3/index.json"
$FeedFlat   = $null   # resolved below from service index
$FeedSearch = $null   # resolved below from service index

# ─────────────────────────────────────────────────────────────────────────────
# Well-known Microsoft AppIds (always added as implicit dependencies)
# ─────────────────────────────────────────────────────────────────────────────
$MsAppIds = [ordered]@{
    #"Application" = "00000000-0000-0000-0000-000000000000"
    #"System"      = "8874ed3a-0643-4247-9ced-7a7002f7135d"
    "LS Central System App" = "2a0547fa-f60f-401d-b1fe-c7c4be8fb230"
    "LS Central" = "5ecfc871-5d82-43f1-9c54-59685e82318d"
}

# ─────────────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────────────
function Log([string]$msg) {
    $ts = [DateTime]::Now.ToString("yyyy-MM-dd HH:mm:ss.ff")
    Write-Host "[$ts] $msg"
}

function Invoke-SafeGet([string]$Uri) {
    try   { return Invoke-RestMethod -Uri $Uri -Method Get -UseBasicParsing -ErrorAction Stop }
    catch { return $null }
}

function Resolve-FeedEndpoints {
    param([string]$IndexUrl)
    $index = Invoke-SafeGet $IndexUrl
    if (-not $index) { throw "Cannot reach NuGet service index: $IndexUrl" }

    $flat   = ($index.resources | Where-Object { $_.'@type' -eq 'PackageBaseAddress/3.0.0'   } | Select-Object -First 1).'@id'
    $search = ($index.resources | Where-Object { $_.'@type' -like 'SearchQueryService*'       } | Select-Object -First 1).'@id'

    if (-not $flat)   { throw "PackageBaseAddress endpoint not found in service index" }
    if (-not $search) { throw "SearchQueryService endpoint not found in service index" }

    # Strip trailing slash so we can append consistently
    return @{
        Flat   = $flat.TrimEnd('/')
        Search = $search.TrimEnd('/')
    }
}

function Find-NuGetPackageId {
    param([string]$AppId)

    # The SearchQueryService on Azure DevOps Artifacts feeds is /query2
    $result = Invoke-SafeGet "$FeedSearch`?q=$AppId&take=300&prerelease=false"
    if (-not $result -or -not $result.data) { return $null }

    # Find package whose ID ends with the AppId (the AL naming convention)
    $candidates = $result.data | Where-Object { $_.id -like "*.$AppId" }
    if (-not $candidates) { return $null }

    # Prefer worldwide (no locale suffix) over localised packages
    $locales = 'da','de','fi','fr','is','it','na','nl','no','nz','ru','se','uk','us',
               'au','be','ca','ch','cz','dk','es','gb','in','mx','my','ph','pl','sg','th','tw','za'
    $localePattern = '\.(' + ($locales -join '|') + ')$'

    $w1 = $candidates | Where-Object { $_.id -notmatch $localePattern } | Select-Object -First 1
    if ($w1) { return $w1.id }
    return ($candidates | Select-Object -First 1).id
}

function Get-BestVersion {
    param([string[]]$Versions, [string]$Required)

    $reqMajor = [int]($Required -split '\.')[0]

    # Exact match
    if ($Required -in $Versions) { return $Required }

    # Highest compatible version with same major
    $compatible = $Versions |
        Where-Object { [int]($_ -split '\.')[0] -eq $reqMajor } |
        Sort-Object  { [version]$_ } -Descending |
        Select-Object -First 1

    return $compatible
}

function Expand-AppFromNupkg {
    param([string]$NupkgPath, [string]$DestFolder)

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::OpenRead($NupkgPath)
    try {
        $appEntry = $zip.Entries | Where-Object { $_.Name -like "*.app" } | Select-Object -First 1
        if (-not $appEntry) { return $null }

        $dest = Join-Path $DestFolder $appEntry.Name
        [System.IO.Compression.ZipFileExtensions]::ExtractToFile($appEntry, $dest, $true)
        return $dest
    }
    finally { $zip.Dispose() }
}

function Download-Symbol {
    param(
        [string]$AppId,
        [string]$AppName,
        [string]$Publisher,
        [string]$RequiredVersion,
        [string]$Destination,
        [bool]  $ForceDownload
    )

    # ── 1. Check cache ────────────────────────────────────────────────────────
    # Escape wildcard chars so the name is safe for -like patterns.
    # Defined here (not inside the if-block) so step 2 can reuse it for fallback.
    $safeName = $AppName -replace '[\[\]\*\?]', '`$0'

    if (-not $ForceDownload) {
        # Match by AppId in filename OR by App display name (covers Microsoft packages
        # like "Microsoft_Application_28.0.0.0.app" where AppId = 00000000-...).
        # Use a word-boundary-style check: require the full name surrounded by _ or end.
        $cached = Get-ChildItem $Destination -Filter "*.app" -ErrorAction SilentlyContinue |
                  Where-Object { $_.Name -like "*$AppId*" -or $_.Name -like "*_${safeName}_*" -or $_.Name -like "${safeName}_*" }
        if ($cached) {
            Log "  -> Already cached: $AppName ($($cached.Name)) — skipping"
            return $true
        }
    }

    # ── 2. Resolve NuGet package ID ───────────────────────────────────────────
    Log "  Resolving package ID for AppId: $AppId, Publisher: $Publisher"
    $pkgId = Find-NuGetPackageId -AppId $AppId
    if (-not $pkgId) {
        # Package not in this feed (typical for Microsoft built-ins like Application/System).
        # Fall back to cache regardless of -Force, so we don't report a false failure.
        $cached = Get-ChildItem $Destination -Filter "*.app" -ErrorAction SilentlyContinue |
                  Where-Object { $_.Name -like "*$AppId*" -or $_.Name -like "*_${safeName}_*" -or $_.Name -like "${safeName}_*" }
        if ($cached) {
            Log "  [INFO] $AppName not found in feed — already cached, skipping"
            return $true
        }
        Log "  [WARN] No package found for AppId $AppId ($AppName) and not in cache"
        return $false
    }
    Log "  Selected package: $pkgId"

    # ── 3. Get available versions ─────────────────────────────────────────────
    $indexUrl = "$FeedFlat/$($pkgId.ToLower())/index.json"
    $verIndex = Invoke-SafeGet $indexUrl
    if (-not $verIndex -or -not $verIndex.versions) {
        Log "  [WARN] Could not retrieve version list for $pkgId"
        return $false
    }

    $selectedVer = Get-BestVersion -Versions $verIndex.versions -Required $RequiredVersion
    if (-not $selectedVer) {
        Log "  [WARN] No compatible version found for $pkgId (required $RequiredVersion)"
        return $false
    }

    if ($selectedVer -ne $RequiredVersion) {
        Log "  Exact version $RequiredVersion not found for $pkgId, using compatible version: $selectedVer"
    }

    # ── 4. Download .nupkg ────────────────────────────────────────────────────
    $pkgIdLow  = $pkgId.ToLower()
    $verLow    = $selectedVer.ToLower()
    $nupkgUrl  = "$FeedFlat/$pkgIdLow/$verLow/$pkgIdLow.$verLow.nupkg"
    $tempFile  = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "$pkgIdLow.$verLow.nupkg")

    try {
        Log "  Downloading $AppName $selectedVer from NuGet feed..."
        Invoke-WebRequest -Uri $nupkgUrl -OutFile $tempFile -UseBasicParsing

        # ── 5. Extract .app ───────────────────────────────────────────────────
        $extracted = Expand-AppFromNupkg -NupkgPath $tempFile -DestFolder $Destination
        if ($extracted) {
            Log "  Downloaded package: $AppName $selectedVer -> $(Split-Path $extracted -Leaf)"
            return $true
        }
        else {
            Log "  [WARN] No .app entry found inside nupkg for $pkgId"
            return $false
        }
    }
    finally {
        if (Test-Path $tempFile) { Remove-Item $tempFile -Force }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────────────────────────────────────

# Locate the AL app.json in the workspace (skip .AL-Go, .git, node_modules)
$workspaceRoot = Split-Path $WorkspacePath -Parent
$appJsonPath   = $null

$candidates = Get-ChildItem -Path $workspaceRoot -Recurse -Filter "app.json" -Depth 3 -ErrorAction SilentlyContinue |
    Where-Object { $_.FullName -notlike "*\.AL-Go\*" -and $_.FullName -notlike "*\.git\*" -and $_.FullName -notlike "*\node_modules\*" }

foreach ($candidate in $candidates) {
    try {
        $parsed = Get-Content $candidate.FullName -Raw | ConvertFrom-Json -ErrorAction Stop
        if ($parsed.PSObject.Properties['platform']) {
            $appJsonPath = $candidate.FullName
            break
        }
    } catch { }
}

if (-not $appJsonPath) { throw "No AL app.json found under workspace: $workspaceRoot" }
Log "Using app.json: $appJsonPath"
$appJson = Get-Content $appJsonPath -Raw | ConvertFrom-Json

# Prepare output folder
$appFolder = Split-Path $appJsonPath -Parent
if (-not $OutputPath) { $OutputPath = Join-Path $appFolder ".alpackages" }
if (-not (Test-Path $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath | Out-Null }

Write-Host "Using reference symbols cache paths: [$OutputPath]"
Log "Using symbols country/region: $Country"

# ── Resolve NuGet endpoints from service index ────────────────────────────────
Log "Resolving NuGet feed endpoints from service index..."
$endpoints  = Resolve-FeedEndpoints -IndexUrl $FeedIndexUrl
$FeedFlat   = $endpoints.Flat
$FeedSearch = $endpoints.Search
Log "  Search : $FeedSearch"
Log "  Flat2  : $FeedFlat"

Log "Starting download of packages to $OutputPath"
Write-Host ""

# Build the work queue: Microsoft implicit + explicit dependencies from app.json
$queue = [System.Collections.Generic.List[hashtable]]::new()

foreach ($entry in $MsAppIds.GetEnumerator()) {
    # Use $appJson.application for all implicit dependencies — it carries the BC major version
    # ($appJson.platform is often left as the placeholder "1.0.0.0" and cannot be used reliably)
    $ver = $appJson.application
    $queue.Add(@{ Name = $entry.Name; Publisher = "Microsoft"; AppId = $entry.Value; Version = $ver })
}

foreach ($dep in $appJson.dependencies) {
    $queue.Add(@{ Name = $dep.name; Publisher = $dep.publisher; AppId = $dep.id; Version = $dep.version })
}

Log "Direct dependencies to download:"
foreach ($item in $queue) {
    Log "  - $($item.Name) $($item.Version) (Publisher: $($item.Publisher), AppId: $($item.AppId))"
}
Write-Host ""

# Download each dependency (skip duplicates)
$seen    = @{}
$success = 0
$failed  = 0

foreach ($item in $queue) {
    if ($seen.ContainsKey($item.AppId)) { continue }
    $seen[$item.AppId] = $true

    $ok = Download-Symbol `
        -AppId          $item.AppId `
        -AppName        $item.Name `
        -Publisher      $item.Publisher `
        -RequiredVersion $item.Version `
        -Destination    $OutputPath `
        -ForceDownload  $Force.IsPresent

    if ($ok) { $success++ } else { $failed++ }
    Write-Host ""
}

# Summary
Log "Completed: $success downloaded, $failed failed."
Log "Total packages available in $OutputPath`: $(@(Get-ChildItem $OutputPath -Filter '*.app').Count)"
