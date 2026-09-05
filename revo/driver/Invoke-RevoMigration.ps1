<#
.SYNOPSIS
    Drives the REVO_Production.xlsm lean-out and rebuild through Excel COM.

.DESCRIPTION
    Everything that mutates the workbook goes through Excel itself. Excel is
    the only thing that deletes a sheet and repairs the VBA document-module
    binding in the same operation - openpyxl, xlsxwriter and every other
    library drop the drawings and the VML form controls that carry the macro
    buttons on REVO Floor, Command Center, Plan Config and Holds.

    The original file is never opened for write. The script copies it, works
    on the copy, and takes a per-phase backup so any single phase can be
    rolled back without unwinding the ones before it.

.PARAMETER Source
    Path to the pristine Revo_Production.xlsm. Never modified.

.PARAMETER OutDir
    Where REVO_Operations.xlsm and the backups are written.

.PARAMETER Phase
    all      - everything (default)
    baseline - copy and inventory only, no mutation
    lean     - import modConfig + modMigrate, run MigrateAll
    layers   - import the remaining modules, build quality + input sheet + form
    verify   - re-open and report, no mutation

.EXAMPLE
    .\Invoke-RevoMigration.ps1 -Source 'C:\REVO\Revo_Production.xlsm' -OutDir 'C:\REVO\out'

.NOTES
    Requires desktop Excel, and for the 'layers' phase:
    File > Options > Trust Center > Trust Center Settings > Macro Settings >
    Trust access to the VBA project object model.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string] $Source,
    [Parameter(Mandatory = $true)][string] $OutDir,
    [ValidateSet('all', 'baseline', 'lean', 'layers', 'verify')][string] $Phase = 'all',
    [string] $VbaDir = (Join-Path $PSScriptRoot '..\vba'),
    [switch] $KeepExcelOpen
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$script:Excel = $null
$script:Book  = $null
$script:Target = Join-Path $OutDir 'REVO_Operations.xlsm'
$script:Transcript = Join-Path $OutDir ('migration-{0:yyyyMMdd-HHmmss}.log' -f (Get-Date))

# xlOpenXMLWorkbookMacroEnabled
$XL_MACRO_ENABLED = 52

function Write-Step {
    param([string] $Message, [string] $Level = 'INFO')
    $line = '{0:HH:mm:ss} [{1,-5}] {2}' -f (Get-Date), $Level, $Message
    switch ($Level) {
        'OK'   { Write-Host $line -ForegroundColor Green }
        'WARN' { Write-Host $line -ForegroundColor Yellow }
        'FAIL' { Write-Host $line -ForegroundColor Red }
        default { Write-Host $line }
    }
    Add-Content -LiteralPath $script:Transcript -Value $line -Encoding UTF8
}

function Assert-Prereqs {
    if (-not (Test-Path -LiteralPath $Source)) { throw "Source not found: $Source" }
    if (-not (Test-Path -LiteralPath $OutDir)) { New-Item -ItemType Directory -Path $OutDir -Force | Out-Null }
    if (-not (Test-Path -LiteralPath $VbaDir)) { throw "VBA module directory not found: $VbaDir" }

    $running = Get-Process -Name EXCEL -ErrorAction SilentlyContinue
    if ($running) {
        Write-Step ("{0} EXCEL.EXE process(es) already running. Close Excel first - an orphaned instance will lock the file." -f $running.Count) 'WARN'
    }
}

# The 'layers' phase imports modules and generates a UserForm. Both need
# programmatic access to the VBA project, which is off by default.
function Test-VbomTrust {
    $found = $false
    $trusted = $false
    Get-ChildItem 'HKCU:\Software\Microsoft\Office' -ErrorAction SilentlyContinue |
        Where-Object { $_.PSChildName -match '^\d+\.\d+$' } |
        ForEach-Object {
            $key = Join-Path $_.PSPath 'Excel\Security'
            if (Test-Path $key) {
                $found = $true
                $v = (Get-ItemProperty -Path $key -Name AccessVBOM -ErrorAction SilentlyContinue).AccessVBOM
                if ($v -eq 1) { $trusted = $true }
            }
        }
    if (-not $found) { Write-Step 'Could not read the Excel Trust Center registry key; continuing.' 'WARN'; return $true }
    return $trusted
}

function New-WorkingCopy {
    Write-Step "Copying source -> $script:Target"
    $srcHashBefore = (Get-FileHash -LiteralPath $Source -Algorithm SHA256).Hash
    Copy-Item -LiteralPath $Source -Destination $script:Target -Force
    $srcHashAfter = (Get-FileHash -LiteralPath $Source -Algorithm SHA256).Hash
    if ($srcHashBefore -ne $srcHashAfter) { throw 'Source file changed during copy. Aborting.' }
    Write-Step ("Source SHA256 {0} - unchanged" -f $srcHashBefore.Substring(0, 16)) 'OK'
    Write-Step ("Working copy: {0:N2} MB" -f ((Get-Item $script:Target).Length / 1MB))
}

function Backup-Phase {
    param([string] $Name)
    $bk = Join-Path $OutDir ("REVO_Operations.before-{0}.xlsm" -f $Name)
    Copy-Item -LiteralPath $script:Target -Destination $bk -Force
    Write-Step "Phase backup: $(Split-Path $bk -Leaf)"
}

function Open-Book {
    Write-Step 'Starting Excel'
    $script:Excel = New-Object -ComObject Excel.Application
    $script:Excel.Visible = $false
    $script:Excel.DisplayAlerts = $false
    $script:Excel.EnableEvents = $false
    $script:Excel.AskToUpdateLinks = $false
    $script:Excel.AutomationSecurity = 3   # msoAutomationSecurityForceDisable - no auto-run macros on open
    $script:Book = $script:Excel.Workbooks.Open($script:Target, 0, $false)
    $script:Excel.AutomationSecurity = 1   # msoAutomationSecurityLow - our own Run calls need macros enabled
    Write-Step ("Opened. {0} sheets." -f $script:Book.Worksheets.Count) 'OK'
}

function Close-Book {
    param([switch] $Save)
    if ($script:Book) {
        if ($Save) {
            Write-Step 'Saving'
            $script:Book.SaveAs($script:Target, $XL_MACRO_ENABLED)
        }
        $script:Book.Close($false)
        $script:Book = $null
    }
}

function Stop-Excel {
    try { if ($script:Book) { $script:Book.Close($false) } } catch { }
    try { if ($script:Excel) { $script:Excel.Quit() } } catch { }
    foreach ($o in @($script:Book, $script:Excel)) {
        if ($o) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } catch { } }
    }
    $script:Book = $null
    $script:Excel = $null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    Write-Step 'Excel released'
}

function Import-Module-Bas {
    param([string[]] $Names)
    foreach ($n in $Names) {
        $path = Join-Path $VbaDir "$n.bas"
        if (-not (Test-Path -LiteralPath $path)) { throw "Module not found: $path" }
        # Replace rather than duplicate: importing twice gives modConfig1.
        $existing = $null
        try { $existing = $script:Book.VBProject.VBComponents.Item($n) } catch { }
        if ($existing) {
            $script:Book.VBProject.VBComponents.Remove($existing)
            Write-Step "Removed existing $n"
        }
        $script:Book.VBProject.VBComponents.Import((Resolve-Path $path).Path) | Out-Null
        Write-Step "Imported $n" 'OK'
    }
}

function Invoke-Macro {
    param([string] $Name, [object[]] $Args = @())
    Write-Step "Run $Name"
    try {
        if ($Args.Count -eq 0) { $script:Book.Application.Run($Name) | Out-Null }
        else { $script:Book.Application.Run($Name, $Args[0]) | Out-Null }
        Write-Step "$Name completed" 'OK'
        return $true
    }
    catch {
        Write-Step "$Name FAILED: $($_.Exception.Message)" 'FAIL'
        return $false
    }
}

function Get-Inventory {
    param([string] $Label)
    Write-Step "--- inventory: $Label ---"
    $rows = @()
    foreach ($ws in $script:Book.Worksheets) {
        $ur = $ws.UsedRange
        $shapes = 0
        try { $shapes = $ws.Shapes.Count } catch { }
        $rows += [pscustomobject]@{
            Sheet   = $ws.Name
            Rows    = $ur.Rows.Count
            Cols    = $ur.Columns.Count
            LastCell= $ur.Address($false, $false)
            Shapes  = $shapes
            Visible = $ws.Visible
        }
    }
    $csv = Join-Path $OutDir ("inventory-{0}.csv" -f $Label)
    $rows | Sort-Object Sheet | Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding UTF8
    Write-Step ("{0} sheets, {1} shapes -> {2}" -f $rows.Count, ($rows | Measure-Object Shapes -Sum).Sum, (Split-Path $csv -Leaf))

    $conns = @()
    foreach ($c in $script:Book.Connections) { $conns += $c.Name }
    Write-Step ("connections: {0}" -f ($(if ($conns) { $conns -join ', ' } else { '(none)' })))
    return $rows
}

# ---------------------------------------------------------------- phases

function Invoke-Baseline {
    New-WorkingCopy
    Open-Book
    $inv = Get-Inventory 'baseline'
    Write-Step ("Baseline: {0} sheets, {1:N2} MB" -f $script:Book.Worksheets.Count, ((Get-Item $script:Target).Length / 1MB)) 'OK'
    Close-Book
    return $inv
}

function Invoke-Lean {
    Backup-Phase 'lean'
    Open-Book
    Import-Module-Bas @('modConfig', 'modMigrate')
    Invoke-Macro 'SetSilent' @($true) | Out-Null
    $ok = Invoke-Macro 'MigrateAll'
    if (-not $ok) { Write-Step 'MigrateAll failed - restore from the before-lean backup.' 'FAIL' }
    Get-Inventory 'after-lean' | Out-Null
    Close-Book -Save
    Write-Step ("After lean: {0:N2} MB" -f ((Get-Item $script:Target).Length / 1MB)) 'OK'
    return $ok
}

function Invoke-Layers {
    if (-not (Test-VbomTrust)) {
        Write-Step 'Trust access to the VBA project object model is OFF.' 'FAIL'
        Write-Step 'File > Options > Trust Center > Trust Center Settings > Macro Settings > tick it, then re-run.' 'FAIL'
        Write-Step 'If policy blocks it, the modules must be imported by hand through the VBE.' 'FAIL'
        return $false
    }
    Backup-Phase 'layers'
    Open-Book
    Import-Module-Bas @('modConfig', 'modQuality', 'modWeeklyUpdate', 'modFormBuilder')
    Invoke-Macro 'SetSilent' @($true) | Out-Null

    $results = [ordered]@{}
    $results['QualityBuild']           = Invoke-Macro 'QualityBuild'
    $results['BuildWeeklyInputSheet']  = Invoke-Macro 'BuildWeeklyInputSheet'
    $results['QualityRefresh']         = Invoke-Macro 'QualityRefresh'      # must survive an empty log
    $results['BuildDispositionForm']   = Invoke-Macro 'BuildDispositionForm'
    $results['RefreshWeeklyMetrics']   = Invoke-Macro 'RefreshWeeklyMetrics'

    foreach ($k in $results.Keys) {
        Write-Step ("{0,-24} {1}" -f $k, $(if ($results[$k]) { 'PASS' } else { 'FAIL' })) $(if ($results[$k]) { 'OK' } else { 'FAIL' })
    }
    Get-Inventory 'after-layers' | Out-Null
    Close-Book -Save
    return (-not ($results.Values -contains $false))
}

function Invoke-Verify {
    Open-Book
    $inv = Get-Inventory 'verify'
    $expect = @(
        'REVO Floor','Plan Config','Plan Engine','Scheduler','Holds','Overrides',
        'Least Resistance Plan','Remaining by Shift','Flow to Plan','Command Center',
        'Weekly Production Plan','Daily Production Plan','Daily Shift Priorities','Production Plan',
        'Capacity Planner','Shift Ramp Plan','Ramp Dashboard','Cummulative Plan vs Release','Release Log',
        'Reject Log','Rework Tracker','Rework From Inventory','Reject_Dashboard','Weekly Reporting',
        'Production Analysis','Summary','Monthly Prod Consolidated',
        'REVO_SALES_FRCST','Upside Production Forecast'
    )
    $have = $inv.Sheet
    $missing = $expect | Where-Object { $_ -notin $have }
    $extra   = $have   | Where-Object { $_ -notin $expect }

    if ($missing) { Write-Step ("MISSING retained sheets: {0}" -f ($missing -join ', ')) 'FAIL' }
    else { Write-Step 'All 29 retained sheets present' 'OK' }
    if ($extra) { Write-Step ("Added sheets (expected: the 4 quality sheets + Weekly_Update_Input): {0}" -f ($extra -join ', ')) }

    $mb = (Get-Item $script:Target).Length / 1MB
    Write-Step ("File size {0:N2} MB (target: under 4 MB)" -f $mb) $(if ($mb -lt 4) { 'OK' } else { 'WARN' })

    $log = Join-Path (Split-Path $script:Target -Parent) 'REVO_Migration_Log.txt'
    if (Test-Path $log) {
        Write-Step '--- REVO_Migration_Log.txt tail ---'
        Get-Content -LiteralPath $log -Tail 40 | ForEach-Object { Write-Step "  $_" }
    }
    Close-Book
    return (-not $missing)
}

# ---------------------------------------------------------------- main

try {
    Write-Step "REVO migration driver. Phase = $Phase"
    Write-Step "Transcript: $script:Transcript"
    Assert-Prereqs

    switch ($Phase) {
        'baseline' { Invoke-Baseline | Out-Null }
        'lean'     { Invoke-Lean     | Out-Null }
        'layers'   { Invoke-Layers   | Out-Null }
        'verify'   { Invoke-Verify   | Out-Null }
        'all' {
            Invoke-Baseline | Out-Null
            if (Invoke-Lean) {
                if (Invoke-Layers) { Invoke-Verify | Out-Null }
                else { Write-Step 'Layers phase failed - stopping before verify.' 'FAIL' }
            }
            else { Write-Step 'Lean phase failed - stopping.' 'FAIL' }
        }
    }
    Write-Step 'Driver finished.' 'OK'
}
catch {
    Write-Step "DRIVER ERROR: $($_.Exception.Message)" 'FAIL'
    Write-Step $_.ScriptStackTrace 'FAIL'
    exit 1
}
finally {
    if (-not $KeepExcelOpen) { Stop-Excel }
    Write-Step "Full transcript: $script:Transcript"
}
