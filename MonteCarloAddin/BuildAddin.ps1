#===============================================================================
# BuildAddin.ps1
# PowerShell build script for MCSimAddin.xlam
# Run: .\BuildAddin.ps1
#===============================================================================

$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$srcDir = Join-Path $scriptDir "src"
$outputPath = Join-Path $scriptDir "MCSimAddin.xlam"

Write-Host "=============================================="  -ForegroundColor Cyan
Write-Host " MCSimAddin Build Script (PowerShell)"          -ForegroundColor Cyan
Write-Host "=============================================="  -ForegroundColor Cyan
Write-Host ""

# Verify source files
if (-not (Test-Path $srcDir)) {
    Write-Host "ERROR: src folder not found" -ForegroundColor Red
    exit 1
}

# Start Excel
Write-Host "Starting Excel..." -ForegroundColor Yellow
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    # Create workbook
    Write-Host "Creating workbook..." -ForegroundColor Yellow
    $wb = $excel.Workbooks.Add()
    $vbProj = $wb.VBProject

    # Import .bas modules
    $modulesDir = Join-Path $srcDir "Modules"
    if (Test-Path $modulesDir) {
        Get-ChildItem -Path $modulesDir -Filter "*.bas" | ForEach-Object {
            Write-Host "  + $($_.Name)" -ForegroundColor Green
            try {
                $vbProj.VBComponents.Import($_.FullName) | Out-Null
            } catch {
                Write-Host "    WARNING: $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }
    }

    # Import .frm forms
    $formsDir = Join-Path $srcDir "Forms"
    if (Test-Path $formsDir) {
        Get-ChildItem -Path $formsDir -Filter "*.frm" | ForEach-Object {
            Write-Host "  + $($_.Name)" -ForegroundColor Green
            try {
                $vbProj.VBComponents.Import($_.FullName) | Out-Null
            } catch {
                Write-Host "    WARNING: $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }
    }

    # Import .cls classes
    $classesDir = Join-Path $srcDir "Classes"
    if (Test-Path $classesDir) {
        Get-ChildItem -Path $classesDir -Filter "*.cls" | ForEach-Object {
            Write-Host "  + $($_.Name)" -ForegroundColor Green
            try {
                $vbProj.VBComponents.Import($_.FullName) | Out-Null
            } catch {
                Write-Host "    WARNING: $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }
    }

    # Update ThisWorkbook
    $twbPath = Join-Path $srcDir "ThisWorkbook.cls"
    if (Test-Path $twbPath) {
        Write-Host "Updating ThisWorkbook..." -ForegroundColor Yellow
        $content = Get-Content $twbPath -Raw
        # Extract code portion
        $codeLines = $content -split "`r?`n"
        $inCode = $false
        $codePart = @()
        foreach ($line in $codeLines) {
            if ($line -match "^(Option|Private|Public)" -or $inCode) {
                $inCode = $true
                $codePart += $line
            }
        }
        $codeText = $codePart -join "`r`n"

        try {
            $twbModule = $vbProj.VBComponents.Item("ThisWorkbook")
            $cm = $twbModule.CodeModule
            if ($cm.CountOfLines -gt 0) {
                $cm.DeleteLines(1, $cm.CountOfLines)
            }
            $cm.AddFromString($codeText)
        } catch {
            Write-Host "    WARNING: Could not update ThisWorkbook" -ForegroundColor Yellow
        }
    }

    # Save as XLAM
    Write-Host ""
    Write-Host "Saving as XLAM..." -ForegroundColor Yellow

    if (Test-Path $outputPath) {
        Remove-Item $outputPath -Force
    }

    # xlOpenXMLAddIn = 55
    $wb.SaveAs($outputPath, 55)
    Write-Host "  Saved: $outputPath" -ForegroundColor Green

    $wb.Close($false)

    Write-Host ""
    Write-Host "=============================================="  -ForegroundColor Green
    Write-Host " Build Complete!"                                -ForegroundColor Green
    Write-Host "=============================================="  -ForegroundColor Green
    Write-Host ""
    Write-Host "Output: $outputPath"
    Write-Host ""
    Write-Host "INSTALLATION:" -ForegroundColor Cyan
    Write-Host "  1. Open Excel"
    Write-Host "  2. File > Options > Add-ins"
    Write-Host "  3. Manage: Excel Add-ins > Go..."
    Write-Host "  4. Browse... > select MCSimAddin.xlam"
    Write-Host "  5. Check the add-in and click OK"
    Write-Host ""
    Write-Host "RIBBON (custom tab):" -ForegroundColor Cyan
    Write-Host "  Use Office RibbonX Editor to inject"
    Write-Host "  src\CustomUI\customUI14.xml into the XLAM"
    Write-Host "  https://github.com/fernandreu/office-ribbonx-editor"
    Write-Host ""

} catch {
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
} finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect()
}
