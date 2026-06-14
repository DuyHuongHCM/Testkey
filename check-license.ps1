$ErrorActionPreference = 'Stop'

function Write-Section {
    param([string]$Title)

    Write-Host ""
    Write-Host "==== $Title ====" -ForegroundColor Cyan
}

function Test-IsWindows {
    return [System.Runtime.InteropServices.RuntimeInformation]::IsOSPlatform(
        [System.Runtime.InteropServices.OSPlatform]::Windows
    )
}

function Get-WindowsLicenseStatus {
    Write-Section "Windows Activation"

    if (-not (Test-IsWindows)) {
        Write-Host "This check requires Windows." -ForegroundColor Yellow
        return
    }

    $slmgr = Join-Path $env:SystemRoot 'System32\slmgr.vbs'
    if (-not (Test-Path $slmgr)) {
        Write-Host "slmgr.vbs not found." -ForegroundColor Red
        return
    }

    Write-Host "Running: slmgr /xpr ..."
    cscript.exe //NoLogo $slmgr /xpr
}

function Get-OfficeScriptCandidates {
    $roots = @($env:ProgramFiles, ${env:ProgramFiles(x86)}) | Where-Object { $_ }
    $candidates = @()

    foreach ($root in $roots) {
        $officeBase = Join-Path $root 'Microsoft Office'

        $commonPath = Join-Path $officeBase 'Office16\OSPP.VBS'
        if (Test-Path $commonPath) {
            $candidates += $commonPath
        }

        $commonPath15 = Join-Path $officeBase 'Office15\OSPP.VBS'
        if (Test-Path $commonPath15) {
            $candidates += $commonPath15
        }

        if (($candidates.Count -eq 0) -and (Test-Path $officeBase)) {
            $dynamic = Get-ChildItem -Path $officeBase -Filter OSPP.VBS -Recurse -ErrorAction SilentlyContinue |
                Select-Object -ExpandProperty FullName
            if ($dynamic) {
                $candidates += $dynamic
            }
        }
    }

    return $candidates | Sort-Object -Unique
}

function Get-OfficeLicenseStatus {
    Write-Section "Office Activation"

    if (-not (Test-IsWindows)) {
        Write-Host "This check requires Windows." -ForegroundColor Yellow
        return
    }

    $scripts = Get-OfficeScriptCandidates
    if (-not $scripts) {
        Write-Host "No OSPP.VBS found. Office may not be installed." -ForegroundColor Yellow
        return
    }

    foreach ($script in $scripts) {
        Write-Host ""
        Write-Host "Checking: $script"
        cscript.exe //NoLogo $script /dstatus
    }
}

Write-Host "Quick license check for Windows + Office" -ForegroundColor Green
Get-WindowsLicenseStatus
Get-OfficeLicenseStatus
