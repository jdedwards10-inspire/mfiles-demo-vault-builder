<#
.SYNOPSIS
  Generate accounting demo data (clients, engagements, documents) as a YAML file
  that Update-MFilesVault.ps1 loads. Writes a REAL file per document - .pdf for
  returns/invoices/statements, .docx for engagement letters, .xlsx for ledgers/
  payroll/workpapers (via DemoFileFactory.ps1, no Office required). Use -PlainText
  for simple .txt, or -NoFiles for metadata-only. Assumes the accounting-firm
  SCHEMA is already applied to the vault.

.EXAMPLE
  # 1) generate
  .\New-DemoData.ps1 -Clients 12 -Documents 50 -ConnectionName "Credit Union"
  # 2) (once) apply the schema, then load the data
  .\Update-MFilesVault.ps1 -YamlPath .\accounting-firm.yaml -ApplySchema -SchemaOnly -Cloud -ConnectionName "Credit Union"
  .\Update-MFilesVault.ps1 -YamlPath .\demo-data.yaml -Cloud -ConnectionName "Credit Union"

.NOTES
  Re-generating + re-loading is safe: objects match by title, so the same set
  converges instead of duplicating. Preparer/Reviewer (Users lookups) are left
  unset so no real vault users are required.
#>
[CmdletBinding()]
param(
    [int]    $Clients = 12,
    [int]    $Documents = 50,
    [int]    $EngagementsPerClient = 2,
    [string] $OutYaml = "demo-data.yaml",
    [string] $FilesDir = "demo-files",
    [string] $ConnectionName,
    [switch] $NoFiles,
    [switch] $PlainText          # use simple .txt files instead of pdf/docx/xlsx
)
$ErrorActionPreference = 'Stop'
Import-Module powershell-yaml -ErrorAction Stop
if (-not $NoFiles -and -not $PlainText) { . "$PSScriptRoot\DemoFileFactory.ps1" }

# ---------- value-list values (must match accounting-firm.yaml) -------------
$fiscalYears   = @('FY2023','FY2024','FY2025','FY2026')
$serviceLines  = @('Tax','Audit & Assurance','Bookkeeping','Advisory','Payroll','Consulting')
$engStatuses   = @('Not Started','In Progress','In Review','Completed','Filed','On Hold')
$filingStatus  = @('Single','Married Filing Jointly','Married Filing Separately','Head of Household')
$entityTypes   = @('Sole Proprietorship','Partnership','LLC','S-Corp','C-Corp','Non-Profit')

# ---------- client name pool ------------------------------------------------
$nameBases = @(
    'Acme Manufacturing','Summit Consulting','Riverside Dental','TechNova Solutions',
    'Green Valley Farms','Harbor Point Realty','Cornerstone Legal','Pinnacle Fitness',
    'Willow Creek Bakery','Metro Logistics','Sunrise Pediatrics','BlueSky Architects',
    'Ironwood Construction','Lakeside Marina','Copperfield Ventures','Evergreen Landscaping',
    'Nova Diagnostics','Redwood Publishing','Silverline Freight','Maple Street Cafe',
    'Quantum Robotics','Heritage Insurance','Brookfield Dairy','Apex Roofing'
)
function Get-EntitySuffix($etype) {
    switch ($etype) {
        'LLC'                 { ' LLC' }
        'S-Corp'              { ' Inc.' }
        'C-Corp'              { ' Inc.' }
        'Partnership'         { ' LLP' }
        'Non-Profit'          { ' Foundation' }
        default               { '' }   # Sole Proprietorship
    }
}
function New-EIN { '{0:00}-{1:0000000}' -f (Get-Random -Max 100), (Get-Random -Max 10000000) }
function FYYear([string]$fy) { [int]($fy -replace '\D','') }
function Rand($arr) { $arr[(Get-Random -Max $arr.Count)] }
function Money { [math]::Round((Get-Random -Minimum 250 -Maximum 75000) + (Get-Random -Max 100)/100.0, 2) }

# ---------- files -----------------------------------------------------------
$filesRoot = $null
if (-not $NoFiles) {
    $filesRoot = (Resolve-Path -LiteralPath (New-Item -ItemType Directory -Force -Path $FilesDir)).Path
}
# document class -> real file format
$fmtByClass = @{
    'Tax Return'          = 'pdf';  'Financial Statement' = 'pdf';  'Invoice' = 'pdf'
    'Receipt'             = 'pdf';  'Bank Statement'      = 'pdf'
    'Engagement Letter'   = 'docx'
    'Payroll Record'      = 'xlsx'; 'Audit Workpaper'     = 'xlsx'; 'General Ledger' = 'xlsx'
}
function New-DemoFile([string]$title, [string]$kind, $props) {
    if ($NoFiles) { return $null }
    $safe = ($title -replace '[\\/:*?"<>|#]', '_').Trim()
    $client = [string]$props['Client']
    $fy     = if ($props.Contains('Fiscal Year')) { [string]$props['Fiscal Year'] } else { '' }
    $amt    = if ($props.Contains('Amount')) { [double]$props['Amount'] } else { [math]::Round((Get-Random -Min 500 -Max 90000) + (Get-Random -Max 100)/100.0, 2) }

    if ($PlainText) {
        $path = Join-Path $filesRoot "$safe.txt"
        $lines = @("$kind", ('=' * $kind.Length), "Title  : $title", "Client : $client")
        foreach ($k in $props.Keys) { if ($k -ne 'Client') { $lines += ('{0,-16}: {1}' -f $k, $props[$k]) } }
        Set-Content -LiteralPath $path -Value $lines -Encoding UTF8
        return $path
    }

    $fmt  = if ($fmtByClass.ContainsKey($kind)) { $fmtByClass[$kind] } else { 'pdf' }
    $path = Join-Path $filesRoot "$safe.$fmt"

    switch ($fmt) {
        'pdf' {
            $lines = @("Client:  $client")
            if ($fy) { $lines += "Fiscal Year:  $fy" }
            switch ($kind) {
                'Invoice'  { $lines += @("Invoice Date:  $($props['Date Received'])", "Due Date:  $($props['Due Date'])", "", "Description                         Amount", "Professional services rendered      `$$('{0:N2}' -f $amt)", "-----------------------------------------", "Total Due:  `$$('{0:N2}' -f $amt)") }
                'Receipt'  { $lines += @("Date:  $($props['Date Received'])", "", "Received with thanks:  `$$('{0:N2}' -f $amt)", "Payment method:  ACH (demo)") }
                'Tax Return' { $lines += @("Tax Year:  $($props['Tax Year'])", "Filing Status:  $($props['Filing Status'])", "", "Adjusted Gross Income:  `$$('{0:N0}' -f (Get-Random -Min 40000 -Max 400000))", "Total Tax:  `$$('{0:N0}' -f (Get-Random -Min 3000 -Max 90000))", "Refund / (Due):  `$$('{0:N0}' -f (Get-Random -Min -8000 -Max 12000))") }
                'Financial Statement' { $ta=Get-Random -Min 100000 -Max 5000000; $tl=Get-Random -Min 20000 -Max 3000000; $lines += @("", "Balance Sheet (summary)", "Total Assets:       `$$('{0:N0}' -f $ta)", "Total Liabilities:  `$$('{0:N0}' -f $tl)", "Total Equity:       `$$('{0:N0}' -f ($ta-$tl))") }
                'Bank Statement' { $ob=Get-Random -Min 5000 -Max 250000; $lines += @("", "Opening Balance:  `$$('{0:N2}' -f $ob)", "Deposits:         `$$('{0:N2}' -f (Get-Random -Min 1000 -Max 90000))", "Withdrawals:      `$$('{0:N2}' -f (Get-Random -Min 500 -Max 70000))", "Closing Balance:  `$$('{0:N2}' -f ($ob + (Get-Random -Min -20000 -Max 40000)))") }
            }
            $lines += @("", "(Demo document for M-Files vault testing.)")
            New-PdfFile -Path $path -Title $kind -Lines $lines
        }
        'docx' {
            $sl = if ($props.Contains('Service Line')) { $props['Service Line'] } else { 'advisory' }
            $lines = @(
                (Get-Date -Format 'MMMM d, yyyy'), "", "Dear $client,", "",
                "We are pleased to confirm the terms of our engagement to provide $sl services" +
                    $(if ($fy) { " for $fy" } else { '' }) + ". This letter outlines the scope of work,",
                "responsibilities, and fee arrangements for the engagement.", "",
                "We appreciate the opportunity to serve you and look forward to working together.", "",
                "Sincerely,", "", "JDE Accounting LLP")
            New-WordFile -Path $path -Title 'Engagement Letter' -Lines $lines
        }
        'xlsx' {
            switch ($kind) {
                'Payroll Record' {
                    $rows = @( ,@('Employee','Gross Pay','Taxes','Net Pay') )
                    foreach ($n in @('A. Nguyen','B. Patel','C. Garcia','D. Smith','E. Okafor')) {
                        $g = Get-Random -Min 3000 -Max 12000; $t = [math]::Round($g*0.28,2)
                        $rows += ,@($n, [double]$g, [double]$t, [double]([math]::Round($g-$t,2)))
                    }
                }
                'General Ledger' {
                    $rows = @( ,@('Account','Debit','Credit') )
                    foreach ($a in @('Cash','Accounts Receivable','Revenue','Rent Expense','Payroll Expense','Supplies')) {
                        $d = Get-Random -Min 0 -Max 50000; $c = if ($d -lt 15000) { Get-Random -Min 0 -Max 50000 } else { 0 }
                        $rows += ,@($a, [double]$d, [double]$c)
                    }
                }
                default {  # Audit Workpaper
                    $rows = @( ,@('Ref','Procedure','Amount','Conclusion') )
                    $j = 1
                    foreach ($p in @('Cash count','AR confirmation','Inventory test','Revenue cutoff','Expense sample')) {
                        $rows += ,@("W-$j", $p, [double](Get-Random -Min 1000 -Max 80000), 'No exceptions')
                        $j++
                    }
                }
            }
            New-ExcelFile -Path $path -Title $kind -Rows $rows
        }
    }
    return $path
}

# ---------- 1) clients ------------------------------------------------------
$clientObjs = @()
$clientList = @()   # {Name; FY-agnostic}
for ($i = 0; $i -lt $Clients; $i++) {
    $base = if ($i -lt $nameBases.Count) { $nameBases[$i] } else { "{0} {1}" -f (Rand $nameBases), ($i + 1) }
    $etype = Rand $entityTypes
    $name  = $base + (Get-EntitySuffix $etype)
    if ($clientList -contains $name) { $name = "$name ($($i+1))" }
    $clientList += $name
    $contact = "$([char](65 + (Get-Random -Max 26))). " + (Rand @('Nguyen','Patel','Garcia','Smith','Kowalski','Okafor','Rossi','Chen','Brooks','Martel'))
    $email   = ('info@' + (($base -replace '[^A-Za-z0-9]','').ToLower()) + '.example')
    $clientObjs += [ordered]@{
        class = 'Client'
        title = $name
        properties = [ordered]@{
            EIN            = New-EIN
            'Entity Type'  = $etype
            'Client Contact' = $contact
            'Client Email'   = $email
            Notes          = "$etype client. Demo record."
        }
    }
}

# ---------- 2) engagements (grouped per client) -----------------------------
$engObjs = @()
$engByClient = @{}
foreach ($name in $clientList) {
    $engByClient[$name] = @()
    $used = @()
    for ($e = 0; $e -lt $EngagementsPerClient; $e++) {
        $fy = Rand $fiscalYears; $sl = Rand $serviceLines
        $key = "$fy|$sl"; if ($used -contains $key) { continue }; $used += $key
        $title = "$name - $fy $sl"
        $engByClient[$name] += $title
        $y = FYYear $fy
        $engObjs += [ordered]@{
            class = 'Engagement'
            title = $title
            properties = [ordered]@{
                Client              = $name
                'Service Line'      = $sl
                'Fiscal Year'       = $fy
                'Engagement Status' = (Rand $engStatuses)
                'Due Date'          = ('{0:0000}-04-15' -f ($y + 1))
                Notes               = "$sl engagement for $fy."
            }
        }
    }
}

# ---------- 3) documents ----------------------------------------------------
# class -> builder that returns the properties hashtable for one document
$docKinds = @('Tax Return','Financial Statement','Invoice','Receipt','Bank Statement','Engagement Letter','Payroll Record','Audit Workpaper','General Ledger')
$docObjs = @()
for ($i = 1; $i -le $Documents; $i++) {
    $clientName = Rand $clientList
    $kind = Rand $docKinds
    $fy = Rand $fiscalYears; $y = FYYear $fy
    $engTitle = if ($engByClient[$clientName].Count) { Rand $engByClient[$clientName] } else { $null }
    $seq = '{0:000}' -f $i
    $props = [ordered]@{ Client = $clientName }
    $title = ''
    switch ($kind) {
        'Tax Return' {
            $title = "Tax Return - $clientName - $fy [$seq]"
            if ($engTitle) { $props['Engagement'] = $engTitle }
            $props['Tax Year']          = $y
            $props['Filing Status']     = (Rand $filingStatus)
            $props['Service Line']      = 'Tax'
            $props['Document Category'] = 'Tax Return'
            $props['Engagement Status'] = (Rand $engStatuses)
            $props['Due Date']          = ('{0:0000}-04-15' -f ($y + 1))
        }
        'Financial Statement' {
            $title = "Financial Statement - $clientName - $fy [$seq]"
            if ($engTitle) { $props['Engagement'] = $engTitle }
            $props['Fiscal Year']       = $fy
            $props['Service Line']      = (Rand $serviceLines)
            $props['Document Category'] = 'Financial Statement'
        }
        'Invoice' {
            $title = "Invoice #$((1000 + $i)) - $clientName"
            $props['Amount']            = (Money)
            $props['Date Received']     = ('{0:0000}-{1:00}-{2:00}' -f $y, (Get-Random -Min 1 -Max 13), (Get-Random -Min 1 -Max 28))
            $props['Due Date']          = ('{0:0000}-{1:00}-{2:00}' -f $y, (Get-Random -Min 1 -Max 13), (Get-Random -Min 1 -Max 28))
            $props['Document Category'] = 'Invoice'
        }
        'Receipt' {
            $title = "Receipt #$((5000 + $i)) - $clientName"
            $props['Amount']            = (Money)
            $props['Date Received']     = ('{0:0000}-{1:00}-{2:00}' -f $y, (Get-Random -Min 1 -Max 13), (Get-Random -Min 1 -Max 28))
            $props['Document Category'] = 'Receipt'
        }
        'Bank Statement' {
            $title = "Bank Statement - $clientName - $fy [$seq]"
            $props['Fiscal Year']       = $fy
            $props['Document Category'] = 'Bank Statement'
        }
        'Engagement Letter' {
            $title = "Engagement Letter - $clientName - $fy [$seq]"
            if ($engTitle) { $props['Engagement'] = $engTitle }
            $props['Service Line']      = (Rand $serviceLines)
            $props['Document Category'] = 'Engagement Letter'
        }
        'Payroll Record' {
            $title = "Payroll Record - $clientName - $fy [$seq]"
            $props['Fiscal Year']       = $fy
            $props['Document Category'] = 'Payroll Record'
        }
        'Audit Workpaper' {
            $title = "Audit Workpaper - $clientName - $fy [$seq]"
            if ($engTitle) { $props['Engagement'] = $engTitle }
            $props['Fiscal Year']       = $fy
            $props['Document Category'] = 'Workpaper'
        }
        'General Ledger' {
            $title = "General Ledger - $clientName - $fy [$seq]"
            $props['Fiscal Year']       = $fy
            $props['Document Category'] = 'General Ledger'
        }
    }
    $obj = [ordered]@{ class = $kind; title = $title; properties = $props }
    $file = New-DemoFile $title $kind $props
    if ($file) { $obj['files'] = @($file) }
    $docObjs += $obj
}

# ---------- emit YAML -------------------------------------------------------
$root = [ordered]@{}
if ($ConnectionName) { $root['connection'] = [ordered]@{ name = $ConnectionName } }
$root['objects'] = @($clientObjs + $engObjs + $docObjs)

$header = @(
    "# Auto-generated accounting demo data ($($clientObjs.Count) clients, $($engObjs.Count) engagements, $($docObjs.Count) documents).",
    "# Load with:  .\Update-MFilesVault.ps1 -YamlPath .\$OutYaml [-Cloud -ConnectionName '<name>']",
    "# Requires the accounting-firm schema to be applied first.",
    ""
) -join "`n"
$yaml = ConvertTo-Yaml $root
Set-Content -LiteralPath $OutYaml -Value ($header + $yaml) -Encoding UTF8

Write-Host ("Wrote {0}: {1} clients, {2} engagements, {3} documents." -f $OutYaml, $clientObjs.Count, $engObjs.Count, $docObjs.Count) -ForegroundColor Green
if (-not $NoFiles) { Write-Host "Placeholder files in: $filesRoot" -ForegroundColor Green }
