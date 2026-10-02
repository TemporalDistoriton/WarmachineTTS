@echo off
setlocal
set "SELF=%~f0"
title TTS Model Backup
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$self=$env:SELF; $raw=[IO.File]::ReadAllText($self); $marker='###PS_PAYLOAD###'; $i=$raw.LastIndexOf($marker); if($i -lt 0){throw 'PowerShell payload marker not found.'}; $code=$raw.Substring($i+$marker.Length); Invoke-Expression $code"
set "ERR=%ERRORLEVEL%"
echo.
if not "%ERR%"=="0" echo Backup finished with errors. Exit code: %ERR%
pause
exit /b %ERR%
###PS_PAYLOAD###

$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$BatPath = $env:SELF
$Root = Split-Path -Parent $BatPath

function Safe-Name([string]$Text, [string]$Fallback = 'Unnamed') {
    if ([string]::IsNullOrWhiteSpace($Text)) { $Text = $Fallback }
    foreach ($c in [IO.Path]::GetInvalidFileNameChars()) {
        $Text = $Text.Replace([string]$c, '_')
    }
    $Text = $Text.Trim().TrimEnd('.')
    if ([string]::IsNullOrWhiteSpace($Text)) { return $Fallback }
    if ($Text.Length -gt 120) { $Text = $Text.Substring(0,120) }
    return $Text
}

function Get-ImageExtension([string]$Path, $Response, [string]$Url) {
    try {
        $uri = [Uri]$Url
        $e = [IO.Path]::GetExtension($uri.AbsolutePath).ToLowerInvariant()
        if ($e -in @('.png','.jpg','.jpeg','.webp','.gif','.bmp','.tif','.tiff')) { return $e }
    } catch {}

    try {
        $ct = [string]$Response.Headers['Content-Type']
        if ($ct) {
            $ct = $ct.Split(';')[0].Trim().ToLowerInvariant()
            switch ($ct) {
                'image/png'  { return '.png' }
                'image/jpeg' { return '.jpg' }
                'image/webp' { return '.webp' }
                'image/gif'  { return '.gif' }
                'image/bmp'  { return '.bmp' }
                'image/tiff' { return '.tif' }
            }
        }
    } catch {}

    try {
        [byte[]]$b = [IO.File]::ReadAllBytes($Path)
        if ($b.Length -ge 8 -and $b[0]-eq 0x89 -and $b[1]-eq 0x50 -and $b[2]-eq 0x4E -and $b[3]-eq 0x47) { return '.png' }
        if ($b.Length -ge 3 -and $b[0]-eq 0xFF -and $b[1]-eq 0xD8 -and $b[2]-eq 0xFF) { return '.jpg' }
        if ($b.Length -ge 6 -and [Text.Encoding]::ASCII.GetString($b,0,6) -match '^GIF8') { return '.gif' }
        if ($b.Length -ge 12 -and [Text.Encoding]::ASCII.GetString($b,0,4) -eq 'RIFF' -and [Text.Encoding]::ASCII.GetString($b,8,4) -eq 'WEBP') { return '.webp' }
    } catch {}
    return '.img'
}

function Download-Asset {
    param(
        [string]$Url,
        [string]$DestinationBase,
        [ValidateSet('Model','Diffuse','Collider')][string]$Kind
    )

    if ([string]::IsNullOrWhiteSpace($Url)) { return @{ Status='Missing URL'; File='' } }

    $temp = $DestinationBase + '.download'
    try {
        $response = Invoke-WebRequest -Uri $Url -OutFile $temp -UseBasicParsing -UserAgent 'Mozilla/5.0 TTS-Model-Backup' -TimeoutSec 90

        if ($Kind -eq 'Diffuse') {
            $ext = Get-ImageExtension -Path $temp -Response $response -Url $Url
        } else {
            $ext = '.obj'
            try {
                $u = [Uri]$Url
                $ue = [IO.Path]::GetExtension($u.AbsolutePath).ToLowerInvariant()
                if ($ue -in @('.obj','.fbx','.dae','.stl','.mesh')) { $ext = $ue }
            } catch {}
        }

        $dest = $DestinationBase + $ext
        Move-Item -LiteralPath $temp -Destination $dest -Force
        return @{ Status='OK'; File=$dest }
    }
    catch {
        if (Test-Path -LiteralPath $temp) { Remove-Item -LiteralPath $temp -Force -ErrorAction SilentlyContinue }
        return @{ Status=('FAILED: ' + $_.Exception.Message); File='' }
    }
}

$jsonFiles = @(Get-ChildItem -LiteralPath $Root -Filter '*.json' -File | Sort-Object Name)
if ($jsonFiles.Count -eq 0) {
    Write-Host "No .json save file was found beside the BAT file:" -ForegroundColor Red
    Write-Host "  $Root"
    exit 2
}

if ($jsonFiles.Count -eq 1) {
    $SaveFile = $jsonFiles[0]
} else {
    Write-Host 'Multiple JSON files found:' -ForegroundColor Cyan
    for ($i=0; $i -lt $jsonFiles.Count; $i++) {
        Write-Host ('  [{0}] {1}' -f ($i+1), $jsonFiles[$i].Name)
    }
    do {
        $choice = Read-Host 'Enter the number of the TTS save to back up'
        $n = 0
        $valid = [int]::TryParse($choice, [ref]$n) -and $n -ge 1 -and $n -le $jsonFiles.Count
    } until ($valid)
    $SaveFile = $jsonFiles[$n-1]
}

Write-Host "`nReading: $($SaveFile.Name)" -ForegroundColor Cyan
try {
    $save = Get-Content -LiteralPath $SaveFile.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
} catch {
    Write-Host "Could not parse the JSON save: $($_.Exception.Message)" -ForegroundColor Red
    exit 3
}

$stamp = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
$saveBase = Safe-Name ([IO.Path]::GetFileNameWithoutExtension($SaveFile.Name)) 'TTS_Save'
$BackupRoot = Join-Path $Root (Join-Path 'TTS Model Backup' ($saveBase + '_' + $stamp))
New-Item -ItemType Directory -Path $BackupRoot -Force | Out-Null

$script:Rows = New-Object System.Collections.Generic.List[object]
$script:UsedNames = @{}
$script:ModelCount = 0
$script:AssetOK = 0
$script:AssetFail = 0

function Unique-BaseName([string]$Folder, [string]$Name, [string]$Guid) {
    $base = Safe-Name $Name $Guid
    $key = ($Folder.ToLowerInvariant() + '|' + $base.ToLowerInvariant())
    if (-not $script:UsedNames.ContainsKey($key)) {
        $script:UsedNames[$key] = 1
        return $base
    }
    # Same visible model name in the same folder: append GUID so nothing is overwritten.
    return (Safe-Name ($base + '_' + $Guid) $Guid)
}

function Backup-CustomModel($Obj, [string[]]$FolderParts) {
    if ($null -eq $Obj.CustomMesh) { return }

    $script:ModelCount++
    $guid = [string]$Obj.GUID
    if ([string]::IsNullOrWhiteSpace($guid)) { $guid = ('NO_GUID_{0:D5}' -f $script:ModelCount) }
    $nickname = [string]$Obj.Nickname

    $folder = $BackupRoot
    foreach ($part in $FolderParts) {
        $folder = Join-Path $folder (Safe-Name $part 'Unnamed Bag')
    }
    New-Item -ItemType Directory -Path $folder -Force | Out-Null

    $baseName = Unique-BaseName -Folder $folder -Name $nickname -Guid $guid
    $meshUrl = [string]$Obj.CustomMesh.MeshURL
    $diffuseUrl = [string]$Obj.CustomMesh.DiffuseURL
    $colliderUrl = [string]$Obj.CustomMesh.ColliderURL

    Write-Host ('[{0}] {1}' -f $script:ModelCount, $baseName) -ForegroundColor White
    Write-Host ('     -> ' + ($FolderParts -join '\')) -ForegroundColor DarkGray

    $model = Download-Asset -Url $meshUrl -DestinationBase (Join-Path $folder ($baseName + '_Model')) -Kind Model
    $diff  = Download-Asset -Url $diffuseUrl -DestinationBase (Join-Path $folder ($baseName + '_Diffuse')) -Kind Diffuse
    $coll  = Download-Asset -Url $colliderUrl -DestinationBase (Join-Path $folder ($baseName + '_Collider')) -Kind Collider

    foreach ($r in @($model,$diff,$coll)) {
        if ($r.Status -eq 'OK') { $script:AssetOK++ }
        elseif ($r.Status -like 'FAILED:*') { $script:AssetFail++ }
    }

    $script:Rows.Add([pscustomobject]@{
        Folder       = ($FolderParts -join '\')
        ModelName    = $nickname
        GUID         = $guid
        FileBase     = $baseName
        MeshURL      = $meshUrl
        MeshStatus   = $model.Status
        DiffuseURL   = $diffuseUrl
        DiffuseStatus= $diff.Status
        ColliderURL  = $colliderUrl
        ColliderStatus=$coll.Status
    }) | Out-Null
}

function Walk-Object($Obj, [string[]]$FolderParts) {
    if ($null -eq $Obj) { return }

    if ([string]$Obj.Name -eq 'Custom_Model') {
        Backup-CustomModel -Obj $Obj -FolderParts $FolderParts
    }

    # Attached/child models remain in the same logical location.
    if ($Obj.PSObject.Properties.Name -contains 'ChildObjects' -and $null -ne $Obj.ChildObjects) {
        foreach ($child in @($Obj.ChildObjects)) {
            Walk-Object -Obj $child -FolderParts $FolderParts
        }
    }

    # Models inside a TTS Bag/Infinite_Bag are saved beneath a folder named after the bag.
    if ($Obj.PSObject.Properties.Name -contains 'ContainedObjects' -and $null -ne $Obj.ContainedObjects) {
        $nextParts = $FolderParts
        $objectType = [string]$Obj.Name
        if ($objectType -match 'Bag') {
            $bagName = [string]$Obj.Nickname
            if ([string]::IsNullOrWhiteSpace($bagName)) { $bagName = [string]$Obj.GUID }
            if ([string]::IsNullOrWhiteSpace($bagName)) { $bagName = 'Unnamed Bag' }
            $safeBagName = Safe-Name $bagName 'Unnamed Bag'
            if ($FolderParts.Count -eq 1 -and $FolderParts[0] -eq 'On the Table') {
                $nextParts = @($safeBagName)
            } else {
                $nextParts = @($FolderParts + $safeBagName)
            }
        }
        foreach ($contained in @($Obj.ContainedObjects)) {
            Walk-Object -Obj $contained -FolderParts $nextParts
        }
    }
}

# Everything top-level starts in the requested "On the Table" folder.
foreach ($obj in @($save.ObjectStates)) {
    Walk-Object -Obj $obj -FolderParts @('On the Table')
}

$manifest = Join-Path $BackupRoot 'backup_manifest.csv'
$script:Rows | Export-Csv -LiteralPath $manifest -NoTypeInformation -Encoding UTF8
Copy-Item -LiteralPath $SaveFile.FullName -Destination (Join-Path $BackupRoot $SaveFile.Name) -Force

Write-Host "`nBackup complete." -ForegroundColor Green
Write-Host "Models found : $script:ModelCount"
Write-Host "Assets saved : $script:AssetOK"
Write-Host "Failed URLs  : $script:AssetFail"
Write-Host "Output       : $BackupRoot" -ForegroundColor Cyan
Write-Host "Manifest     : $manifest"

if ($script:AssetFail -gt 0) { exit 1 } else { exit 0 }
