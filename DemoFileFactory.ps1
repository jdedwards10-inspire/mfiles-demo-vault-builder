<#
  DemoFileFactory.ps1 - dot-sourced helpers that create REAL .pdf / .docx / .xlsx
  files with no Microsoft Office and no external libraries (pure .NET).
    New-PdfFile   -Path <p> -Title <t> -Lines <string[]>
    New-WordFile  -Path <p> -Title <t> -Lines <string[]>
    New-ExcelFile -Path <p> -Title <t> -Rows <object[][]>   # first row = header
#>

Add-Type -AssemblyName System.IO.Compression      -ErrorAction SilentlyContinue
Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue

function ConvertTo-XmlText { param([string]$s) [System.Security.SecurityElement]::Escape([string]$s) }

# --- PDF: minimal one-page document with a proper cross-reference table -----
function New-PdfFile {
    param([string]$Path, [string]$Title, [string[]]$Lines)
    $latin1 = [System.Text.Encoding]::GetEncoding('ISO-8859-1')   # 1 byte per char => char length == byte length
    function esc([string]$s) {
        $s = ($s -replace '[^\x20-\x7E]', ' ')                    # drop non-ASCII (Helvetica/WinAnsi safe)
        $s -replace '\\', '\\\\' -replace '\(', '\(' -replace '\)', '\)'
    }
    $c  = "BT`n/F1 18 Tf`n72 740 Td`n(" + (esc $Title) + ") Tj`n/F1 11 Tf`n0 -28 Td`n"
    foreach ($ln in $Lines) { $c += "(" + (esc $ln) + ") Tj`n0 -16 Td`n" }
    $c += "ET"

    $objs = @(
        "<< /Type /Catalog /Pages 2 0 R >>",
        "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>",
        "<< /Length $($c.Length) >>`nstream`n$c`nendstream",
        "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>"
    )
    $pdf = "%PDF-1.4`n"
    $offsets = @()
    for ($i = 0; $i -lt $objs.Count; $i++) {
        $offsets += $pdf.Length
        $pdf += "$($i + 1) 0 obj`n$($objs[$i])`nendobj`n"
    }
    $xref = $pdf.Length
    $pdf += "xref`n0 $($objs.Count + 1)`n0000000000 65535 f `n"
    foreach ($o in $offsets) { $pdf += ('{0:0000000000} 00000 n ' -f $o) + "`n" }
    $pdf += "trailer`n<< /Size $($objs.Count + 1) /Root 1 0 R >>`nstartxref`n$xref`n%%EOF"
    [System.IO.File]::WriteAllText($Path, $pdf, $latin1)
}

# --- ZIP (OOXML) helpers ----------------------------------------------------
function Add-ZipText {
    param($Zip, [string]$Name, [string]$Content)
    $entry = $Zip.CreateEntry($Name, [System.IO.Compression.CompressionLevel]::Optimal)
    $s = $entry.Open()
    try { $b = [System.Text.Encoding]::UTF8.GetBytes($Content); $s.Write($b, 0, $b.Length) } finally { $s.Dispose() }
}
function New-OoxmlZip {
    param([string]$Path, [hashtable]$Parts)   # name -> xml string
    if (Test-Path -LiteralPath $Path) { Remove-Item -LiteralPath $Path -Force }
    $zip = [System.IO.Compression.ZipFile]::Open($Path, [System.IO.Compression.ZipArchiveMode]::Create)
    try { foreach ($k in $Parts.Keys) { Add-ZipText $zip $k $Parts[$k] } } finally { $zip.Dispose() }
}

# --- DOCX -------------------------------------------------------------------
function New-WordFile {
    param([string]$Path, [string]$Title, [string[]]$Lines)
    $body = "<w:p><w:r><w:rPr><w:b/><w:sz w:val=""32""/></w:rPr><w:t xml:space=""preserve"">$(ConvertTo-XmlText $Title)</w:t></w:r></w:p>"
    foreach ($ln in $Lines) { $body += "<w:p><w:r><w:t xml:space=""preserve"">$(ConvertTo-XmlText $ln)</w:t></w:r></w:p>" }
    $parts = @{
        '[Content_Types].xml' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
        '_rels/.rels'         = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
        'word/document.xml'   = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' + $body + '<w:sectPr/></w:body></w:document>'
    }
    New-OoxmlZip -Path $Path -Parts $parts
}

# --- XLSX -------------------------------------------------------------------
function Get-ColLetter { param([int]$Index) $s=''; $n=$Index+1; while ($n -gt 0) { $m=($n-1)%26; $s=[char](65+$m)+$s; $n=[int](($n-1)/26) }; $s }
function New-ExcelFile {
    param([string]$Path, [string]$Title, [object[][]]$Rows)
    $sd = ''
    for ($r = 0; $r -lt $Rows.Count; $r++) {
        $cells = ''
        $row = $Rows[$r]
        for ($cIdx = 0; $cIdx -lt $row.Count; $cIdx++) {
            $ref = (Get-ColLetter $cIdx) + ($r + 1)
            $val = $row[$cIdx]
            if ($val -is [int] -or $val -is [long] -or $val -is [double] -or $val -is [decimal]) {
                $cells += "<c r=""$ref""><v>$val</v></c>"
            } else {
                $cells += "<c r=""$ref"" t=""inlineStr""><is><t xml:space=""preserve"">$(ConvertTo-XmlText ([string]$val))</t></is></c>"
            }
        }
        $sd += "<row r=""$($r+1)"">$cells</row>"
    }
    $parts = @{
        '[Content_Types].xml' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>'
        '_rels/.rels'         = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
        'xl/workbook.xml'     = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>'
        'xl/_rels/workbook.xml.rels' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>'
        'xl/worksheets/sheet1.xml'   = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>' + $sd + '</sheetData></worksheet>'
    }
    New-OoxmlZip -Path $Path -Parts $parts
}
