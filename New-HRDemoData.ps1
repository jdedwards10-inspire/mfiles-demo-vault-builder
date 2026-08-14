<#
.SYNOPSIS
  Generate HR demo data (departments, employees, documents) as a YAML file that
  Update-MFilesVault.ps1 loads. Writes a REAL file per document (.pdf/.docx/.xlsx
  via DemoFileFactory.ps1, no Office required). Assumes the hr-vault SCHEMA is
  already applied.

.EXAMPLE
  .\New-HRDemoData.ps1 -Departments 6 -Employees 15 -Documents 50 -ConnectionName "JDE-HR"
  .\Update-MFilesVault.ps1 -YamlPath .\hr-demo-data.yaml -Cloud -ConnectionName "JDE-HR"

.NOTES
  Manager / Department Head (Employee lookups) are left unset to avoid ordering
  issues; everything else is populated. Re-loading the same YAML is idempotent.
#>
[CmdletBinding()]
param(
    [int]    $Departments = 6,
    [int]    $Employees = 15,
    [int]    $Documents = 50,
    [string] $OutYaml = "hr-demo-data.yaml",
    [string] $FilesDir = "hr-demo-files",
    [string] $ConnectionName,
    [switch] $NoFiles,
    [switch] $PlainText
)
$ErrorActionPreference = 'Stop'
Import-Module powershell-yaml -ErrorAction Stop
if (-not $NoFiles -and -not $PlainText) { . "$PSScriptRoot\DemoFileFactory.ps1" }

# ---------- value-list values (must match hr-vault.yaml) --------------------
$empTypes  = @('Full-Time','Part-Time','Contractor','Intern','Temporary')
$empStatus = @('Active','On Leave','Probation','Terminated','Retired')
$jobLevels = @('Entry','Associate','Senior','Lead','Manager','Director','VP','Executive')
$locations = @('New York','San Francisco','Chicago','Austin','Remote','London','Toronto')
$ratings   = @('Exceeds Expectations','Meets Expectations','Needs Improvement','Unsatisfactory')
$years     = @('2023','2024','2025','2026')

$deptNames = @('Human Resources','Finance','Engineering','Sales','Marketing','Operations','Legal','Information Technology','Customer Support','Executive')
$titlesByLevel = @{
    'Entry'='Associate'; 'Associate'='Specialist'; 'Senior'='Senior Specialist'; 'Lead'='Team Lead'
    'Manager'='Manager'; 'Director'='Director'; 'VP'='Vice President'; 'Executive'='Chief Officer'
}
$firstNames = @('James','Mary','John','Patricia','Robert','Jennifer','Michael','Linda','David','Elizabeth','Maria','Ahmed','Wei','Priya','Carlos','Sofia','Noah','Emma','Liam','Olivia','Kenji','Aisha','Diego','Fatima')
$lastNames  = @('Smith','Johnson','Nguyen','Patel','Garcia','Kowalski','Okafor','Rossi','Chen','Brooks','Martel','Andersson','Haddad','Ivanov','Suzuki','Silva','Murphy','Reyes','Kim','Dubois')

function Rand($arr) { $arr[(Get-Random -Max $arr.Count)] }
function Money { [math]::Round((Get-Random -Minimum 42000 -Maximum 240000) / 100.0, 0) * 100 }
function DateIn([string]$y) { '{0}-{1:00}-{2:00}' -f $y, (Get-Random -Min 1 -Max 13), (Get-Random -Min 1 -Max 28) }

$filesRoot = $null
if (-not $NoFiles) { $filesRoot = (Resolve-Path -LiteralPath (New-Item -ItemType Directory -Force -Path $FilesDir)).Path }

# document class -> real file format
$fmtByClass = @{
    'Offer Letter'='docx'; 'Disciplinary Notice'='docx'; 'Policy Acknowledgment'='docx'
    'Employment Contract'='pdf'; 'Resume'='pdf'; 'Performance Review'='pdf'
    'Tax Form'='pdf'; 'Training Certificate'='pdf'
    'Payslip'='xlsx'
}
function New-HRFile([string]$title, [string]$kind, $props) {
    if ($NoFiles) { return $null }
    $safe = ($title -replace '[\\/:*?"<>|#]', '_').Trim()
    $emp  = [string]$props['Employee']
    $dept = if ($props.Contains('Department')) { [string]$props['Department'] } else { '' }
    if ($PlainText) {
        $path = Join-Path $filesRoot "$safe.txt"
        $lines = @("$kind", ('=' * $kind.Length), "Title   : $title", "Employee: $emp")
        foreach ($k in $props.Keys) { if ($k -ne 'Employee') { $lines += ('{0,-18}: {1}' -f $k, $props[$k]) } }
        Set-Content -LiteralPath $path -Value $lines -Encoding UTF8
        return $path
    }
    $fmt  = if ($fmtByClass.ContainsKey($kind)) { $fmtByClass[$kind] } else { 'pdf' }
    $path = Join-Path $filesRoot "$safe.$fmt"
    switch ($fmt) {
        'pdf' {
            $lines = @("Employee:  $emp")
            if ($dept) { $lines += "Department:  $dept" }
            switch ($kind) {
                'Employment Contract' { $lines += @("Employment Type:  $($props['Employment Type'])", "Effective Date:  $($props['Effective Date'])", "", "This agreement sets out the terms of employment between the Company and the", "employee named above, including duties, compensation, and confidentiality.") }
                'Resume'              { $lines += @("", "PROFILE", "Experienced professional seeking to contribute to the organization.", "", "EXPERIENCE", " - Prior Role, Previous Company (3 yrs)", " - Earlier Role, Another Company (2 yrs)", "", "EDUCATION", " - B.S., State University") }
                'Performance Review'  { $lines += @("Review Year:  $($props['Year'])", "Rating:  $($props['Review Rating'])", "", "Summary: The employee met or exceeded objectives for the review period.", "Goals for next year established and agreed.") }
                'Tax Form'            { $lines += @("Tax Year:  $($props['Year'])", "", "Form W-4 (demo) - Employee's Withholding Certificate", "Filing status and allowances on record.") }
                'Training Certificate'{ $lines += @("Effective Date:  $($props['Effective Date'])", "", "CERTIFICATE OF COMPLETION", "This certifies the employee has completed required compliance training.") }
            }
            $lines += @("", "(Demo document for M-Files vault testing.)")
            New-PdfFile -Path $path -Title $kind -Lines $lines
        }
        'docx' {
            switch ($kind) {
                'Offer Letter' {
                    $lines = @((Get-Date -Format 'MMMM d, yyyy'), "", "Dear $emp,", "",
                        "We are delighted to offer you the position of $($props['Job Title']) in the" +
                            $(if ($dept) { " $dept department" } else { '' }) + ", based in $($props['Location']).",
                        "We believe your skills will be a great addition to our team.", "",
                        "Please review the enclosed terms and confirm your acceptance.", "",
                        "Warm regards,", "", "People Operations")
                }
                'Disciplinary Notice' {
                    $lines = @((Get-Date -Format 'MMMM d, yyyy'), "", "RE: Formal Notice", "",
                        "Dear $emp,", "", "This letter serves as a formal notice regarding a workplace policy matter.",
                        "A meeting has been scheduled to discuss the details and next steps.", "",
                        "Human Resources")
                }
                default {  # Policy Acknowledgment
                    $lines = @("EMPLOYEE POLICY ACKNOWLEDGMENT", "",
                        "Employee: $emp", "Effective Date: $($props['Effective Date'])", "",
                        "I acknowledge that I have received, read, and understood the Company's",
                        "policies, including the code of conduct and information-security policy.", "",
                        "Signature: __________________________")
                }
            }
            New-WordFile -Path $path -Title $kind -Lines $lines
        }
        'xlsx' {  # Payslip
            $gross = [double]$props['Salary'] / 24.0
            $tax   = [math]::Round($gross * 0.22, 2); $ins = [math]::Round($gross * 0.06, 2)
            $ret   = [math]::Round($gross * 0.05, 2); $net = [math]::Round($gross - $tax - $ins - $ret, 2)
            $rows = @(
                ,@('Payslip', $emp)
                ,@('Pay Date', [string]$props['Pay Date'])
                ,@('', '')
                ,@('Item','Amount')
                ,@('Gross Pay', [double][math]::Round($gross,2))
                ,@('Federal Tax', [double](-$tax))
                ,@('Health Insurance', [double](-$ins))
                ,@('Retirement 401(k)', [double](-$ret))
                ,@('Net Pay', [double]$net)
            )
            New-ExcelFile -Path $path -Title 'Payslip' -Rows $rows
        }
    }
    return $path
}

# ---------- 1) departments --------------------------------------------------
$deptObjs = @(); $deptList = @()
for ($i = 0; $i -lt $Departments; $i++) {
    $name = if ($i -lt $deptNames.Count) { $deptNames[$i] } else { "Department $($i+1)" }
    if ($deptList -contains $name) { $name = "$name ($($i+1))" }
    $deptList += $name
    $deptObjs += [ordered]@{
        class = 'Department'
        title = $name
        properties = [ordered]@{
            'Cost Center' = 'CC-{0:000}' -f (100 + $i)
            Location      = (Rand $locations)
            Notes         = "$name department."
        }
    }
}

# ---------- 2) employees ----------------------------------------------------
$empObjs = @(); $empList = @(); $empByDept = @{}
foreach ($d in $deptList) { $empByDept[$d] = @() }
for ($i = 0; $i -lt $Employees; $i++) {
    $fn = Rand $firstNames; $ln = Rand $lastNames
    $name = "$fn $ln"
    if ($empList -contains $name) { $name = "$fn $ln $($i+1)" }
    $empList += $name
    $dept  = Rand $deptList; $empByDept[$dept] += $name
    $level = Rand $jobLevels
    $title = $titlesByLevel[$level]
    $empObjs += [ordered]@{
        class = 'Employee'
        title = $name
        properties = [ordered]@{
            'Employee ID'       = 'E{0:00000}' -f (10000 + $i)
            Department          = $dept
            'Job Title'         = $title
            'Employment Type'   = (Rand $empTypes)
            'Employment Status' = (Rand $empStatus)
            'Job Level'         = $level
            Location            = (Rand $locations)
            'Work Email'        = ('{0}.{1}@democo.example' -f $fn.ToLower(), $ln.ToLower())
            Phone               = '(555) {0:000}-{1:0000}' -f (Get-Random -Max 1000), (Get-Random -Max 10000)
            'Hire Date'         = (DateIn (Rand $years))
            Salary              = (Money)
            Notes               = "$level $title in $dept."
        }
    }
}

# ---------- 3) documents ----------------------------------------------------
$docKinds = @('Offer Letter','Employment Contract','Resume','Performance Review','Payslip','Tax Form','Training Certificate','Disciplinary Notice','Policy Acknowledgment')
$docObjs = @()
for ($i = 1; $i -le $Documents; $i++) {
    $emp  = Rand $empList
    $dept = $null
    foreach ($d in $deptList) { if ($empByDept[$d] -contains $emp) { $dept = $d; break } }
    $kind = Rand $docKinds
    $yr   = Rand $years
    $seq  = '{0:000}' -f $i
    $props = [ordered]@{ Employee = $emp }
    $title = ''
    switch ($kind) {
        'Offer Letter' {
            $title = "Offer Letter - $emp [$seq]"
            if ($dept) { $props['Department'] = $dept }
            $props['Job Title']         = (Rand @('Analyst','Engineer','Manager','Specialist','Coordinator','Director'))
            $props['Job Level']         = (Rand $jobLevels)
            $props['Location']          = (Rand $locations)
            $props['Document Category'] = 'Offer Letter'
            $props['Effective Date']    = (DateIn $yr)
        }
        'Employment Contract' {
            $title = "Employment Contract - $emp [$seq]"
            if ($dept) { $props['Department'] = $dept }
            $props['Employment Type']   = (Rand $empTypes)
            $props['Document Category'] = 'Employment Contract'
            $props['Effective Date']    = (DateIn $yr)
        }
        'Resume' {
            $title = "Resume - $emp [$seq]"
            $props['Document Category'] = 'Resume'
        }
        'Performance Review' {
            $title = "Performance Review - $emp - $yr [$seq]"
            if ($dept) { $props['Department'] = $dept }
            $props['Year']              = $yr
            $props['Review Rating']     = (Rand $ratings)
            $props['Document Category'] = 'Performance Review'
        }
        'Payslip' {
            $title = "Payslip - $emp - $yr [$seq]"
            $props['Salary']            = [double](Money)
            $props['Pay Date']          = (DateIn $yr)
            $props['Year']              = $yr
            $props['Document Category'] = 'Payslip'
        }
        'Tax Form' {
            $title = "Tax Form W-4 - $emp - $yr [$seq]"
            $props['Year']              = $yr
            $props['Document Category'] = 'Tax Form'
        }
        'Training Certificate' {
            $title = "Training Certificate - $emp [$seq]"
            $props['Document Category'] = 'Training Certificate'
            $props['Effective Date']    = (DateIn $yr)
        }
        'Disciplinary Notice' {
            $title = "Disciplinary Notice - $emp [$seq]"
            if ($dept) { $props['Department'] = $dept }
            $props['Document Category'] = 'Disciplinary Notice'
            $props['Effective Date']    = (DateIn $yr)
        }
        'Policy Acknowledgment' {
            $title = "Policy Acknowledgment - $emp [$seq]"
            $props['Document Category'] = 'Policy Acknowledgment'
            $props['Effective Date']    = (DateIn $yr)
        }
    }
    $obj = [ordered]@{ class = $kind; title = $title; properties = $props }
    $file = New-HRFile $title $kind $props
    if ($file) { $obj['files'] = @($file) }
    $docObjs += $obj
}

# ---------- emit YAML -------------------------------------------------------
$root = [ordered]@{}
if ($ConnectionName) { $root['connection'] = [ordered]@{ name = $ConnectionName } }
$root['objects'] = @($deptObjs + $empObjs + $docObjs)

$header = @(
    "# Auto-generated HR demo data ($($deptObjs.Count) departments, $($empObjs.Count) employees, $($docObjs.Count) documents).",
    "# Load with:  .\Update-MFilesVault.ps1 -YamlPath .\$OutYaml [-Cloud -ConnectionName '<name>']",
    "# Requires the hr-vault schema to be applied first.",
    ""
) -join "`n"
Set-Content -LiteralPath $OutYaml -Value ($header + (ConvertTo-Yaml $root)) -Encoding UTF8

Write-Host ("Wrote {0}: {1} departments, {2} employees, {3} documents." -f $OutYaml, $deptObjs.Count, $empObjs.Count, $docObjs.Count) -ForegroundColor Green
if (-not $NoFiles) { Write-Host "Files in: $filesRoot" -ForegroundColor Green }
