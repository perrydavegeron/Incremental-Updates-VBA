# JIRA Survey Status Automation v2.1
# Developed by Perry Dave V. Geron
# Run on Windows with Microsoft Excel desktop installed.
# Requires temporary "Trust access to the VBA project object model".
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Read-ZipEntryText {
    param($zip, [string] $name)
    $entry = $zip.GetEntry($name)
    if ($null -eq $entry) { throw "Open XML package component not found: $name" }
    $reader = New-Object System.IO.StreamReader($entry.Open())
    try { $reader.ReadToEnd() } finally { $reader.Dispose() }
}
function Set-ZipEntryText {
    param($zip, [string] $name, [string] $content)
    $old = $zip.GetEntry($name)
    if ($null -ne $old) { $old.Delete() }
    $new = $zip.CreateEntry($name)
    $writer = [System.IO.StreamWriter]::new($new.Open(), [System.Text.UTF8Encoding]::new($false))
    try { $writer.Write($content) } finally { $writer.Dispose() }
}
function Add-JiraRibbon {
    param([string] $addInPath)
    $ribbonXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
  <ribbon>
    <tabs>
      <tab id="jiraSurveyStatusTab" label="Survey Status">
        <group id="jiraSurveyStatusGroup" label="JIRA survey tools">
          <button id="jiraSurveyStatusButton" label="Generate Survey Status"
            size="large" imageMso="FileOpen" onAction="JSS_RibbonRun"
            screentip="Generate Survey Status"
            supertip="Create a new survey status report from a JIRA export." />
          <button id="jiraSurveyUpdateButton" label="Update Survey Status"
            size="large" imageMso="RefreshAll" onAction="JSS_UpdateRibbon"
            screentip="Update Survey Status"
            supertip="Append new JIRA tickets and refresh existing records while keeping stakeholder notes." />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
'@
    $stream = [System.IO.File]::Open($addInPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite)
    $zip = [System.IO.Compression.ZipArchive]::new($stream, [System.IO.Compression.ZipArchiveMode]::Update, $false)
    try {
        Set-ZipEntryText $zip 'customUI/customUI14.xml' $ribbonXml
        [xml] $relationships = Read-ZipEntryText $zip '_rels/.rels'
        $ns = 'http://schemas.openxmlformats.org/package/2006/relationships'
        $link = $relationships.CreateElement('Relationship', $ns)
        $link.SetAttribute('Id', 'rIdJiraSurveyRibbon')
        $link.SetAttribute('Type', 'http://schemas.microsoft.com/office/2007/relationships/ui/extensibility')
        $link.SetAttribute('Target', 'customUI/customUI14.xml')
        [void] $relationships.DocumentElement.AppendChild($link)
        Set-ZipEntryText $zip '_rels/.rels' $relationships.OuterXml
        [xml] $contentTypes = Read-ZipEntryText $zip '[Content_Types].xml'
        $tns = 'http://schemas.openxmlformats.org/package/2006/content-types'
        $override = $contentTypes.CreateElement('Override', $tns)
        $override.SetAttribute('PartName', '/customUI/customUI14.xml')
        $override.SetAttribute('ContentType', 'application/vnd.ms-office.customUI+xml')
        [void] $contentTypes.DocumentElement.AppendChild($override)
        Set-ZipEntryText $zip '[Content_Types].xml' $contentTypes.OuterXml
    } finally { $zip.Dispose() }
}

$excel = $null
$buildBook = $null
$registrationBook = $null
$target = $null
try {
    $source = Join-Path $PSScriptRoot 'source\modJiraSurveyStatus.bas'
    if (!(Test-Path -LiteralPath $source)) {
        throw 'Source module is missing. Extract the full ZIP before installing.'
    }
    if (Get-Process -Name EXCEL -ErrorAction SilentlyContinue) {
        throw 'Close every Excel window and save your work before running this installer.'
    }
    $targetDir = Join-Path $env:APPDATA 'Microsoft\AddIns'
    [void](New-Item -ItemType Directory -Force -Path $targetDir)
    $target = Join-Path $targetDir 'JIRA Survey Status Automation.xlam'
    Write-Host "Building the add-in with desktop Excel..." -ForegroundColor Cyan
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $buildBook = $excel.Workbooks.Add()
    try {
        $vbProject = $buildBook.VBProject
        $null = $vbProject.VBComponents.Count
    } catch {
        throw "VBA project access is blocked. In Excel, enable File > Options > Trust Center > Trust Center Settings > Macro Settings > Trust access to the VBA project object model, then retry. Details: $_"
    }
    # Excel VBA imports .bas modules using the Windows legacy text encoding.
    # A UTF-8 BOM may appear as the literal characters 'ï»¿' before Attribute VB_Name,
    # causing the VBA 'Expected: expression' compile error. Remove that risk by
    # creating a BOM-free ASCII source file for the VBComponents.Import operation.
    $importPath = Join-Path $env:TEMP ('JiraSurveyStatus_' + [guid]::NewGuid().ToString('N') + '.bas')
    try {
        $moduleText = [System.IO.File]::ReadAllText($source)
        $moduleText = $moduleText.TrimStart([char]0xFEFF)
        if ($moduleText -notmatch '^Attribute VB_Name\s*=') {
            throw 'The VBA source does not begin with its expected module attribute.'
        }
        if ($moduleText -match '[^\x00-\x7F]') {
            throw 'The VBA source contains non-ASCII characters and needs an ANSI-safe conversion.'
        }
        [System.IO.File]::WriteAllText($importPath, $moduleText, [System.Text.Encoding]::ASCII)
        $module = $vbProject.VBComponents.Import($importPath)
        $module.Name = 'modJiraSurveyStatus'
    } finally {
        if (Test-Path -LiteralPath $importPath) { Remove-Item -LiteralPath $importPath -Force }
    }
    try {
        $buildBook.BuiltinDocumentProperties.Item('Title').Value = 'JIRA Survey Status Automation'
        $buildBook.BuiltinDocumentProperties.Item('Author').Value = 'Perry Dave V. Geron'
    } catch {
        Write-Host 'Workbook metadata could not be set; continuing.' -ForegroundColor Yellow
    }
    if (Test-Path -LiteralPath $target) {
        # Keep the currently working add-in until the replacement is ready.
        Write-Host 'Existing add-in detected; protecting it during the build.' -ForegroundColor Cyan
    }
    # Excel itself compiles and saves the VBA project to a genuine XLAM.
    $staged = Join-Path $targetDir 'JIRA Survey Status Automation_v2_staging.xlam'
    if (Test-Path -LiteralPath $staged) { Remove-Item -LiteralPath $staged -Force }
    $buildBook.SaveAs($staged, 55)
    $buildBook.Close($false)
    $buildBook = $null
    $excel.Quit()
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
    $excel = $null
    Write-Host 'Adding the Excel Ribbon button...' -ForegroundColor Cyan
    Add-JiraRibbon $staged
    # Only after a genuine, ribbon-enabled XLAM exists do we replace v1.
    # Preserve the previous working version for easy recovery.
    $previous = Join-Path $targetDir 'JIRA Survey Status Automation_previous.xlam'
    if (Test-Path -LiteralPath $previous) { Remove-Item -LiteralPath $previous -Force }
    if (Test-Path -LiteralPath $target) { Move-Item -LiteralPath $target -Destination $previous -Force }
    try {
        Move-Item -LiteralPath $staged -Destination $target -Force
    } catch {
        # The old add-in is preserved in case replacement fails midway.
        if (!(Test-Path -LiteralPath $target) -and (Test-Path -LiteralPath $previous)) {
            Move-Item -LiteralPath $previous -Destination $target -Force
        }
        throw
    }

    # Excel.AddIns.Add may fail on some installations, so registration is
    # best effort; the generated XLAM can always be installed manually.
    Write-Host 'Attempting to enable the add-in in Excel...' -ForegroundColor Cyan
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $registrationBook = $excel.Workbooks.Add()
        $addIn = $null
        for ($i = 1; $i -le $excel.AddIns.Count; $i++) {
            $candidate = $excel.AddIns.Item($i)
            if ([string]::Equals([string]$candidate.FullName, [string]$target,
                     [System.StringComparison]::OrdinalIgnoreCase)) {
                $addIn = $candidate
                break
            }
        }
        if ($null -eq $addIn) { $addIn = $excel.AddIns.Add($target, $false) }
        $addIn.Installed = $true
        Write-Host 'The add-in was registered successfully.' -ForegroundColor Green
    } catch {
        Write-Host 'Automatic registration was not available on this Excel installation.' -ForegroundColor Yellow
        Write-Host "Install it manually: Excel > File > Options > Add-ins > Excel Add-ins > Go > Browse: $target" -ForegroundColor Yellow
    } finally {
        if ($null -ne $registrationBook) { try { $registrationBook.Close($false) } catch {} }
        if ($null -ne $excel) { try { $excel.Quit() } catch {} }
        $registrationBook = $null
        $excel = $null
    }
    Write-Host ''
    Write-Host 'BUILD COMPLETE: JIRA Survey Status Automation v2.1' -ForegroundColor Green
    Write-Host "Add-in location: $target"
    Write-Host 'Excel > Survey Status tab > Generate Survey Status / Update Survey Status.'
    Write-Host 'Turn off VBA project model access again after installation.'
} catch {
    Write-Host ''
    Write-Host "INSTALLATION ERROR: $($_.Exception.Message)" -ForegroundColor Red
    if ($target -and (Test-Path -LiteralPath $target)) {
        Write-Host "An XLAM may have been generated at: $target" -ForegroundColor Yellow
    }
    exit 1
} finally {
    if ($null -ne $buildBook) { try { $buildBook.Close($false) } catch {} }
    if ($null -ne $excel) { try { $excel.Quit() } catch {} }
    [GC]::Collect()
}
