<#
.SYNOPSIS
    Creates a formatted .docx file listing a directory structure as a table of contents.
.DESCRIPTION
    Recursively enumerates all subdirectories (files are omitted) under the given
    root and writes them into a Word document as a table of contents: one indented
    entry per directory, children directly below their parents, siblings sorted
    case-insensitively by name.

    Layout:
      Default format:
        - Centered title, root path and generation stamp, separated by a rule
        - Level 1 entries bold 12 pt, level 2 11 pt, level 3+ 10 pt
        - 0.25 inch indent per level
      Simple format (-Format Simple):
        - Plain Consolas 10 pt throughout, no bold, no size changes
        - 2-space indent per level, single line spacing, no extra gaps -
        renders like a plain text editor listing

    The .docx is written directly as an Open XML zip package via
    System.IO.Compression, so no Word installation or COM automation is required
    and there is no limit on the number of entries. Directories that cannot be
    accessed are skipped. Runs on Windows PowerShell 5.1 and PowerShell 7+.
.PARAMETER Directory
    Root directory to list. Defaults to the current working directory.
.PARAMETER OutputFile
    Path of the output .docx file. Relative paths are resolved against the current
    working directory. Defaults to .\DirectoryStructure.docx.
.PARAMETER Title
    Document title. Defaults to "Table of Contents".
.PARAMETER Format
    Output layout: "Default" (styled table of contents) or "Simple" (plain
    monospace listing with space indentation and no extra line spacing).
    Defaults to "Default".
.EXAMPLE
    .\DirectoryToc.ps1
    Writes .\DirectoryStructure.docx for the current directory.
.EXAMPLE
    .\DirectoryToc.ps1 -Directory "C:\Projects\GV2026" -OutputFile "C:\temp\GV2026.docx"
.EXAMPLE
    .\DirectoryToc.ps1 . .\Structure.docx -Title "GV2026 Structure"
.EXAMPLE
    .\DirectoryToc.ps1 -Directory "C:\Projects" -Format Simple
    Writes a plain, text-editor-like listing.
.NOTES
    Exit codes: 0 = success, 1 = invalid input or write error.
#>

[CmdletBinding()]
param(
    [Parameter(Position = 0, HelpMessage = "Root directory to list")]
    [string]$Directory = (Get-Location).ProviderPath,

    [Parameter(Position = 1, HelpMessage = "Output .docx file path")]
    [string]$OutputFile = "DirectoryStructure.docx",

    [Parameter(HelpMessage = "Document title")]
    [string]$Title = "Table of Contents",

    [Parameter(Position = 2, HelpMessage = "Output layout: Default or Simple")]
    [ValidateSet('Default', 'Simple')]
    [string]$Format = 'Default'
)

# --- Helpers -------------------------------------------------------------------

function Esc([string]$s) {
    $s.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;').Replace('"', '&quot;')
}

function Write-ZipEntryText($Zip, [string]$EntryName, [string]$Xml) {
    $entry = $Zip.CreateEntry($EntryName)
    $stream = $entry.Open()
    try {
        $writer = New-Object System.IO.StreamWriter($stream,
            (New-Object System.Text.UTF8Encoding($false)))
        try { $writer.Write($Xml) } finally { $writer.Dispose() }
    }
    finally { $stream.Dispose() }
}

function Fail([string]$Message) {
    Write-Error -Message $Message -ErrorAction Continue
    exit 1
}

function Write-Para($Writer, [string]$Text, [string]$Rpr, [string]$Ppr) {
    # Emits one <w:p> paragraph; empty $Text yields a text-less paragraph.
    $p = '<w:p>'
    if ($Ppr)  { $p += '<w:pPr>' + $Ppr + '</w:pPr>' }
    if ($Text) {
        $p += '<w:r>'
        if ($Rpr) { $p += '<w:rPr>' + $Rpr + '</w:rPr>' }
        $p += '<w:t xml:space="preserve">' + (Esc $Text) + '</w:t></w:r>'
    }
    $Writer.Write($p + '</w:p>')
}

# --- Validate and normalize input ----------------------------------------------

$cwd = (Get-Location).ProviderPath
if ([string]::IsNullOrEmpty($cwd)) { $cwd = (Get-Location).Path }

if ([string]::IsNullOrWhiteSpace($Directory)) { $Directory = $cwd }
if (-not (Test-Path -LiteralPath $Directory -PathType Container)) {
    Fail "Directory not found: '$Directory'"
}
$Directory = (Resolve-Path -LiteralPath $Directory).ProviderPath

# Root path without trailing slash (except drive roots such as 'C:\')
$rootNorm = $Directory.TrimEnd('\')
if ($rootNorm -eq '') { $rootNorm = '\' }
if ($rootNorm -match '^[A-Za-z]:$') { $rootNorm += '\' }

if ([string]::IsNullOrWhiteSpace($Title)) { $Title = "Table of Contents" }

$outLeaf = Split-Path -Path $OutputFile -Leaf
$outDir  = Split-Path -Path $OutputFile -Parent
if ([string]::IsNullOrEmpty($outDir)) {
    $OutputFile = Join-Path $cwd $outLeaf
}
else {
    if (-not (Test-Path -LiteralPath $outDir -PathType Container)) {
        Fail "Output directory not found: '$outDir'"
    }
    $OutputFile = Join-Path (Resolve-Path -LiteralPath $outDir).ProviderPath $outLeaf
}
if ([System.IO.Path]::GetExtension($OutputFile) -ne '.docx') {
    Write-Warning "Output file does not have a .docx extension: '$OutputFile'"
}

Write-Host "Directory : $Directory" -ForegroundColor Cyan
Write-Host "Output    : $OutputFile" -ForegroundColor Cyan

# Probe catches output files that cannot be written *before* the slow scan, so
# users do not wait minutes only to learn the file is open in Word.
if (Test-Path -LiteralPath $OutputFile -PathType Container) {
    Fail "Output path is a directory: '$OutputFile'"
}
if (Test-Path -LiteralPath $OutputFile) {
    try {
        # Clear ReadOnly first, mirroring what Remove-Item -Force would do later,
        # so a read-only file is not a false alarm here.
        $attr = Get-Item -LiteralPath $OutputFile -Force
        if ($attr -is [System.IO.FileInfo] -and $attr.IsReadOnly) { $attr.IsReadOnly = $false }
        $probeStream = [System.IO.File]::Open($OutputFile,
            [System.IO.FileMode]::Open, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
        $probeStream.Dispose()
    }
    catch { Fail "Cannot write '$OutputFile'. Is the file open in another program? ($($_.Exception.Message))" }
}

# --- Scan ----------------------------------------------------------------------

Write-Host "Scanning..." -ForegroundColor Cyan
$allDirs = @(Get-ChildItem -LiteralPath $Directory -Directory -Recurse -ErrorAction SilentlyContinue)
$dirNoun = if ($allDirs.Count -eq 1) { 'directory' } else { 'directories' }
Write-Host "Found $($allDirs.Count) $dirNoun." -ForegroundColor Cyan

# --- Order depth-first: children follow parents, siblings sorted by name --------

$childrenOf = @{}
foreach ($d in $allDirs) {
    $parentPath = if ($d.Parent) { $d.Parent.FullName } else { $rootNorm }
    if (-not $childrenOf.ContainsKey($parentPath)) { $childrenOf[$parentPath] = [System.Collections.Generic.List[object]]::new() }
    $childrenOf[$parentPath].Add($d)
}
foreach ($key in @($childrenOf.Keys)) {
    $childrenOf[$key] = @($childrenOf[$key] | Sort-Object -Property Name -Descending)
}

$ordered = [System.Collections.Generic.List[object]]::new()
$stack   = [System.Collections.Generic.Stack[object]]::new()
foreach ($d in $childrenOf[$rootNorm]) { $stack.Push($d) }
while ($stack.Count -gt 0) {
    $d = $stack.Pop()
    $ordered.Add($d)
    foreach ($kid in $childrenOf[$d.FullName]) { $stack.Push($kid) }
}
# Safety net: append anything the tree walk could not reach
if ($ordered.Count -lt $allDirs.Count) {
    $seen = @{}
    foreach ($d in $ordered) { $seen[$d.FullName] = $true }
    foreach ($d in ($allDirs | Sort-Object -Property FullName)) {
        if (-not $seen.ContainsKey($d.FullName)) { $ordered.Add($d) }
    }
}

# --- Build the document header ---------------------------------------------------

$isSimple = ($Format -eq 'Simple')
if ($isSimple) {
    $plainRpr = '<w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/><w:sz w:val="20"/><w:szCs w:val="20"/>'
    $plainPpr = '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
}

# Function emitting the header paragraphs, so both formats share one code path.
function Write-Header($Writer) {
    if ($script:isSimple) {
        Write-Para $Writer $script:Title  $script:plainRpr $script:plainPpr
        Write-Para $Writer $script:Directory $script:plainRpr $script:plainPpr
        Write-Para $Writer ("Generated {0} - {1} {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm'), $script:ordered.Count, $script:dirNoun) $script:plainRpr $script:plainPpr
        Write-Para $Writer '' $null $script:plainPpr
    }
    else {
        Write-Para $Writer $script:Title ('<w:b/><w:sz w:val="36"/><w:szCs w:val="36"/>') ('<w:jc w:val="center"/><w:spacing w:after="120"/>')
        Write-Para $Writer $script:Directory ('<w:color w:val="595959"/><w:sz w:val="20"/><w:szCs w:val="20"/>') ('<w:jc w:val="center"/><w:spacing w:after="60"/>')
        Write-Para $Writer ("Generated {0} - {1} {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm'), $script:ordered.Count, $script:dirNoun) ('<w:i/><w:color w:val="808080"/><w:sz w:val="18"/><w:szCs w:val="18"/>') ('<w:jc w:val="center"/><w:spacing w:after="200"/>')
        Write-Para $Writer '' $null ('<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="BFBFBF"/></w:pBdr><w:spacing w:after="240"/>')
    }
}

if ($isSimple) {
    $defFonts = '<w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:eastAsia="Consolas" w:cs="Consolas"/>'
    $defSize  = '<w:sz w:val="20"/><w:szCs w:val="20"/>'
}
else {
    $defFonts = '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Times New Roman"/>'
    $defSize  = '<w:sz w:val="22"/><w:szCs w:val="22"/>'
}

$stylesXml = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    '<w:docDefaults>' +
        '<w:rPrDefault><w:rPr>' +
            $defFonts +
            $defSize + '<w:lang w:val="en-US"/>' +
        '</w:rPr></w:rPrDefault>' +
        '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>' +
    '</w:docDefaults>' +
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">' +
        '<w:name w:val="Normal"/><w:qFormat/>' +
    '</w:style>' +
    '</w:styles>'
)

# --- Write the .docx package (no Word/COM required) -----------------------------
# Written as a plain Open XML zip package. System.IO.Packaging is avoided on
# purpose: on .NET Framework it buffers parts in isolated storage and fails on
# large documents with "Unable to determine the identity of domain".

Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem

$ctMain    = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'
$ctStyles  = 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml'
$ctCore    = 'application/vnd.openxmlformats-package.core-properties+xml'
$relOffice = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'
$relStyles = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles'
$relCore   = 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties'

$contentTypesXml = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="' + $ctMain + '"/>' +
    '<Override PartName="/word/styles.xml" ContentType="' + $ctStyles + '"/>' +
    '<Override PartName="/docProps/core.xml" ContentType="' + $ctCore + '"/>' +
    '</Types>'
)

$rootRelsXml = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="' + $relOffice + '" Target="word/document.xml"/>' +
    '<Relationship Id="rId2" Type="' + $relCore + '" Target="docProps/core.xml"/>' +
    '</Relationships>'
)

$docRelsXml = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="' + $relStyles + '" Target="styles.xml"/>' +
    '</Relationships>'
)

$stamp = [DateTime]::UtcNow.ToString("yyyy-MM-dd'T'HH:mm:ss'Z'")
$coreXml = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"' +
    ' xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/"' +
    ' xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
    '<dc:title>' + (Esc $Title) + '</dc:title>' +
    '<dc:creator>DirectoryToc.ps1</dc:creator>' +
    '<cp:lastModifiedBy>DirectoryToc.ps1</cp:lastModifiedBy>' +
    '<dcterms:created xsi:type="dcterms:W3CDTF">' + $stamp + '</dcterms:created>' +
    '<dcterms:modified xsi:type="dcterms:W3CDTF">' + $stamp + '</dcterms:modified>' +
    '</cp:coreProperties>'
)

if (Test-Path -LiteralPath $OutputFile) {
    try { Remove-Item -LiteralPath $OutputFile -Force -ErrorAction Stop }
    catch { Fail "Cannot overwrite '$OutputFile'. Is the file open in another program? ($($_.Exception.Message))" }
}

Write-Host "Writing..." -ForegroundColor Cyan
$fileStream = $null
$zip = $null
try {
    $fileStream = New-Object System.IO.FileStream($OutputFile,
        [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
    $zip = New-Object System.IO.Compression.ZipArchive($fileStream,
        [System.IO.Compression.ZipArchiveMode]::Create)

    Write-ZipEntryText $zip '[Content_Types].xml'          $contentTypesXml
    Write-ZipEntryText $zip '_rels/.rels'                  $rootRelsXml
    Write-ZipEntryText $zip 'docProps/core.xml'            $coreXml
    Write-ZipEntryText $zip 'word/_rels/document.xml.rels' $docRelsXml
    Write-ZipEntryText $zip 'word/styles.xml'              $stylesXml

    # document.xml is streamed paragraph by paragraph so huge trees stay
    # memory-friendly.
    $docEntry  = $zip.CreateEntry('word/document.xml')
    $docStream = $docEntry.Open()
    $docWriter = New-Object System.IO.StreamWriter($docStream,
        (New-Object System.Text.UTF8Encoding($false)), 65536)
    try {
        $docWriter.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
        $docWriter.Write('<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>')
        Write-Header $docWriter

        if ($ordered.Count -eq 0) {
            if ($isSimple) {
                Write-Para $docWriter '(no subdirectories)' $plainRpr $plainPpr
            }
            else {
                Write-Para $docWriter '(no subdirectories)' ('<w:i/><w:color w:val="808080"/>') ('<w:spacing w:before="120"/>')
            }
        }
        else {
            $indentStep = 360   # twips per level (0.25 inch)
            foreach ($d in $ordered) {
                $rel   = $d.FullName.Substring([Math]::Max(0, $rootNorm.Length)).TrimStart('\')
                $depth = [Math]::Max(1, $rel.Split('\').Count)
                if ($isSimple) {
                    Write-Para $docWriter (('  ' * ($depth - 1)) + $d.Name) $plainRpr $plainPpr
                }
                else {
                    $indent = ($depth - 1) * $indentStep
                    if ($depth -le 1) {
                        $rpr = '<w:b/><w:sz w:val="24"/><w:szCs w:val="24"/>'; $before = '160'
                    }
                    elseif ($depth -eq 2) {
                        $rpr = '<w:sz w:val="22"/><w:szCs w:val="22"/>'; $before = '40'
                    }
                    else {
                        $rpr = '<w:sz w:val="20"/><w:szCs w:val="20"/>'; $before = '20'
                    }
                    Write-Para $docWriter $d.Name $rpr ('<w:spacing w:before="' + $before + '" w:after="0"/><w:ind w:left="' + $indent + '"/>')
                }
            }
        }

        $docWriter.Write(
            '<w:sectPr>' +
            '<w:pgSz w:w="12240" w:h="15840"/>' +
            '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
            '</w:sectPr></w:body></w:document>')
    }
    finally { $docWriter.Dispose() }

    $zip.Dispose()
    $zip = $null
    $fileStream.Dispose()
    $fileStream = $null
}
catch {
    if ($zip) { try { $zip.Dispose() } catch { } }
    if ($fileStream) { try { $fileStream.Dispose() } catch { } }
    if (Test-Path -LiteralPath $OutputFile) {
        Remove-Item -LiteralPath $OutputFile -Force -ErrorAction SilentlyContinue
    }
    Fail "Failed to write '$OutputFile': $($_.Exception.Message)"
}

Write-Host ""
Write-Host "Done! Saved to: $OutputFile" -ForegroundColor Green
exit 0




