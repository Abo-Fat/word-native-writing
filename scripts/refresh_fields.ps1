param(
    [Parameter(Mandatory = $true)][string]$InputPath,
    [Parameter(Mandatory = $true)][string]$OutputPath,
    [string]$PythonCommand = 'python'
)

$ErrorActionPreference = 'Stop'
$inputFull = (Resolve-Path -LiteralPath $InputPath).Path
$outputFull = [IO.Path]::GetFullPath($OutputPath)
if ([string]::Equals($inputFull, $outputFull, [StringComparison]::OrdinalIgnoreCase)) {
    throw 'Input and output must be different files.'
}
if (Test-Path -LiteralPath $outputFull) { throw "Refusing to overwrite: $outputFull" }
$parent = [IO.Path]::GetDirectoryName($outputFull)
if (-not (Test-Path -LiteralPath $parent)) { throw "Output directory must exist: $parent" }
$statePath = $outputFull + '.flatopc.xml'
if (Test-Path -LiteralPath $statePath) { throw "State file already exists: $statePath" }
Add-Type -AssemblyName System.IO.Compression.FileSystem
$package = [IO.Compression.ZipFile]::OpenRead($inputFull)
try {
    $entry = $package.GetEntry('word/settings.xml')
    if ($entry) {
        $stream = $entry.Open()
        try {
            $settingsXml = New-Object System.Xml.XmlDocument
            $settingsXml.Load($stream)
        } finally { $stream.Dispose() }
        $manager = New-Object System.Xml.XmlNamespaceManager($settingsXml.NameTable)
        $manager.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
        $update = $settingsXml.SelectSingleNode('/w:settings/w:updateFields', $manager)
        if ($update -and $update.GetAttribute('val', $manager.LookupNamespace('w')) -notin @('false', '0', 'off')) {
            throw 'Disable updateFields in a working copy before opening it for selective refresh.'
        }
    }
} finally { $package.Dispose() }
$word = $null
$doc = $null
$savedOptions = @{}

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $word.AutomationSecurity = 3
    $word.ScreenUpdating = $false
    foreach ($name in @('UpdateLinksAtOpen', 'UpdateFieldsAtPrint', 'UpdateLinksAtPrint', 'Pagination')) {
        $savedOptions[$name] = $word.Options.$name
        $word.Options.$name = $false
    }
    $doc = $word.Documents.Open($inputFull, $false, $true)
    $doc.ActiveWindow.View.Type = 1
    Write-Output ("Opened read-only: {0}; automation HWND={1}" -f $doc.Name, $word.Hwnd)
    $count = $doc.Fields.Count
    foreach ($types in @(@(10, 12), @(3))) {
        for ($i = 1; $i -le $count; $i++) {
            $field = $doc.Fields.Item($i)
            if ($field.Type -in $types) { [void]$field.Update() }
        }
    }
    $doc.Repaginate()
    for ($pass = 0; $pass -lt 2; $pass++) {
        for ($i = 1; $i -le $count; $i++) {
            $field = $doc.Fields.Item($i)
            if ($field.Type -eq 37) { [void]$field.Update() }
        }
        foreach ($section in $doc.Sections) {
            foreach ($footer in $section.Footers) { [void]$footer.Range.Fields.Update() }
        }
        $doc.Repaginate()
    }
    [IO.File]::WriteAllText($statePath, $doc.WordOpenXML, [Text.UTF8Encoding]::new($false))
    Write-Output ("Refreshed fields; Word page count={0}; preserved state={1}" -f $doc.ComputeStatistics(2), $statePath)
} finally {
    if ($null -ne $doc) {
        $doc.Close(0)
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($doc)
    }
    if ($null -ne $word) {
        foreach ($name in $savedOptions.Keys) { $word.Options.$name = $savedOptions[$name] }
        if ($word.Documents.Count -eq 0) { $word.Quit() }
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

& $PythonCommand -X utf8 (Join-Path $PSScriptRoot 'package_word_state.py') $statePath $outputFull
if ($LASTEXITCODE -ne 0) { throw 'Flat OPC packaging failed; the XML state has been retained.' }
Write-Output "Created: $outputFull"
