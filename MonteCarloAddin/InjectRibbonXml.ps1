#===============================================================================
# InjectRibbonXml.ps1
# Injects customUI14.xml into the XLAM file (which is a ZIP archive).
# This is an OPTIONAL step - the toolbar works without it.
# The ribbon XML adds a native "MC Simulation" tab to the Excel ribbon.
#
# Usage: .\InjectRibbonXml.ps1 [-XlamPath "MCSimAddin.xlam"]
#===============================================================================
param(
    [string]$XlamPath = (Join-Path $PSScriptRoot "MCSimAddin.xlam")
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.IO.Compression.FileSystem

$cuiXmlPath = Join-Path $PSScriptRoot "src\CustomUI\customUI14.xml"

if (-not (Test-Path $XlamPath)) {
    Write-Host "ERROR: XLAM not found at $XlamPath" -ForegroundColor Red
    Write-Host "Build the XLAM first with BuildAddin.ps1" -ForegroundColor Yellow
    exit 1
}
if (-not (Test-Path $cuiXmlPath)) {
    Write-Host "ERROR: customUI14.xml not found at $cuiXmlPath" -ForegroundColor Red
    exit 1
}

Write-Host "Injecting ribbon XML into $XlamPath ..." -ForegroundColor Cyan

# Read the customUI XML content
$cuiContent = [System.IO.File]::ReadAllBytes($cuiXmlPath)

# Open the XLAM as a ZIP archive
$zip = [System.IO.Compression.ZipFile]::Open($XlamPath, [System.IO.Compression.ZipArchiveMode]::Update)

try {
    # Remove existing customUI entries if present
    $existing = $zip.Entries | Where-Object { $_.FullName -like "customUI/*" -or $_.FullName -eq "customUI14/customUI14.xml" }
    foreach ($entry in $existing) {
        $entry.Delete()
    }

    # Add customUI14/customUI14.xml
    $newEntry = $zip.CreateEntry("customUI/customUI14.xml", [System.IO.Compression.CompressionLevel]::Optimal)
    $stream = $newEntry.Open()
    $stream.Write($cuiContent, 0, $cuiContent.Length)
    $stream.Close()

    # Update [Content_Types].xml to register the customUI part
    $ctEntry = $zip.GetEntry("[Content_Types].xml")
    if ($ctEntry) {
        $reader = New-Object System.IO.StreamReader($ctEntry.Open())
        $ctXml = $reader.ReadToEnd()
        $reader.Close()

        # Check if customUI content type already exists
        if ($ctXml -notmatch "customUI") {
            # Add the content type for customUI
            $insertion = '<Override PartName="/customUI/customUI14.xml" ContentType="application/xml" />'
            $ctXml = $ctXml -replace '</Types>', "$insertion`n</Types>"

            # Rewrite [Content_Types].xml
            $ctEntry.Delete()
            $newCt = $zip.CreateEntry("[Content_Types].xml", [System.IO.Compression.CompressionLevel]::Optimal)
            $writer = New-Object System.IO.StreamWriter($newCt.Open())
            $writer.Write($ctXml)
            $writer.Close()
        }
    }

    # Add/update .rels to reference the customUI part
    $relsEntry = $zip.GetEntry("_rels/.rels")
    if ($relsEntry) {
        $reader = New-Object System.IO.StreamReader($relsEntry.Open())
        $relsXml = $reader.ReadToEnd()
        $reader.Close()

        if ($relsXml -notmatch "customUI") {
            $relInsertion = '<Relationship Id="rCustomUI" Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" Target="customUI/customUI14.xml" />'
            $relsXml = $relsXml -replace '</Relationships>', "$relInsertion`n</Relationships>"

            $relsEntry.Delete()
            $newRels = $zip.CreateEntry("_rels/.rels", [System.IO.Compression.CompressionLevel]::Optimal)
            $writer = New-Object System.IO.StreamWriter($newRels.Open())
            $writer.Write($relsXml)
            $writer.Close()
        }
    }

    Write-Host "SUCCESS: Ribbon XML injected." -ForegroundColor Green
    Write-Host "The XLAM now has a native 'MC Simulation' ribbon tab." -ForegroundColor Green

} finally {
    $zip.Dispose()
}
