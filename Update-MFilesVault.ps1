<#
.SYNOPSIS
    Build an M-Files (on-prem) vault - structure and objects - from a YAML file.

.DESCRIPTION
    One tool, two internal paths:
      * Schema (-ApplySchema): value lists, object types, property definitions and
        classes, created over an ADMINISTRATIVE (Windows-integrated) connection.
        Alias-driven and idempotent; object-type "users can create objects" is
        enabled automatically; auto-created lookup props and default classes are
        adopted (not duplicated).
      * Objects: the 'objects:' records, created as a named M-Files user (-User, or
        prompted). All COM access is LATE-BOUND against the installed M-Files client,
        so no interop assembly / DLL is required (works regardless of server version).

.PARAMETER YamlPath        Path to the YAML file (except with -ListConnections).
.PARAMETER ConnectionName  Registered connection name (overrides connection.name).
.PARAMETER VaultGuid       Vault GUID (overrides connection.vault).
.PARAMETER ListConnections Print registered connections and exit.
.PARAMETER ListSchema      (alias -ListClasses) Connect read-only and print the vault's
                           object types, classes (+their properties), property
                           definitions, and value lists, then exit. Use it to discover
                           the real names when loading/tagging into an existing vault.
.PARAMETER ApplySchema     Create/refresh the structure sections.
.PARAMETER SchemaPhase     Limit schema apply (all|valuelists|objecttypes|properties|classes).
.PARAMETER SchemaOnly      Apply schema only; do not create objects.
.PARAMETER User            M-Files login to connect as. If omitted, you are prompted
                           for it (defaulting to the connection's stored user, if any).
.PARAMETER Password        SecureString; if omitted you are prompted at runtime.
.PARAMETER NetAddress/Port/Protocol  Server endpoint for the object connection.
.PARAMETER CreateMissingLookups      Add value-list items referenced by objects if missing.
.PARAMETER SingleLogin     Use one M-Files-user connection (-User) for BOTH schema and
                           objects, instead of a separate Windows-integrated admin
                           connection for schema. The user must have admin rights.
.PARAMETER UseConnectionEndpoint  (alias -Cloud) Connect using the registered
                           connection's own endpoint (protocol/address/port/encryption)
                           instead of localhost - required for a CLOUD vault. Implies
                           single-login as -User; the account needs "Full control of
                           vault" (vault admin), NOT server system-administrator.
.PARAMETER AuthType        M-Files auth type: 1=current Windows, 2=specific Windows,
                           3=specific M-Files user (default).
.PARAMETER Encrypted       Force an encrypted (TLS) connection.
.PARAMETER DryRun          Validate only; write nothing.

.EXAMPLE
    .\Update-MFilesVault.ps1 -YamlPath .\accounting-firm.yaml -ApplySchema
.EXAMPLE
    .\Update-MFilesVault.ps1 -YamlPath .\accounting-firm.yaml -ApplySchema -SchemaOnly
#>
[CmdletBinding()]
param(
    [string] $YamlPath,
    [string] $ConnectionName,
    [string] $VaultGuid,
    [switch] $ListConnections,
    [Alias('ListClasses')]
    [switch] $ListSchema,
    [switch] $ApplySchema,
    [ValidateSet('all','valuelists','objecttypes','properties','classes','workflows','views')]
    [string] $SchemaPhase = 'all',
    [switch] $SchemaOnly,
    [string] $User,
    [System.Security.SecureString] $Password,
    [string] $NetAddress = 'localhost',
    [string] $Port = '2266',
    [string] $Protocol = 'ncacn_ip_tcp',
    [Alias('Cloud')]
    [switch] $UseConnectionEndpoint,
    [switch] $Encrypted,
    [int]    $AuthType = 3,
    [switch] $CreateMissingLookups,
    [switch] $AllowDuplicates,
    [Alias('Local','OnPrem')]
    [switch] $SingleLogin,
    [switch] $DryRun
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# --------------------------------------------------------------------------
# Constants
# --------------------------------------------------------------------------
$MFDatatypeText = 1; $MFDatatypeInteger = 2; $MFDatatypeFloating = 3
$MFDatatypeDate = 5; $MFDatatypeTime = 6; $MFDatatypeTimestamp = 7
$MFDatatypeBoolean = 8; $MFDatatypeLookup = 9; $MFDatatypeMultiSelectLookup = 10
$MFDatatypeMultiLineText = 13

$MFBuiltInPropertyDefNameOrTitle = 0
$MFBuiltInPropertyDefClass       = 100
$MFBuiltInPropertyDefWorkflow    = 38
$MFBuiltInPropertyDefState       = 39
$MFBuiltInValueListUsers         = 17
$MFAuthTypeSpecificMFilesUser    = 3

$SchemaTypeMap = @{
    text = 1; multiline = 13; integer = 2; number = 3; date = 5
    time = 6; timestamp = 7; boolean = 8; lookup = 9; multilookup = 10
}

# --------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------
function Write-Info  { param($m) Write-Host "[i] $m" -ForegroundColor Cyan }
function Write-Ok    { param($m) Write-Host "[+] $m" -ForegroundColor Green }
function Write-Warn2 { param($m) Write-Host "[!] $m" -ForegroundColor Yellow }

function ConvertTo-Array {
    param($Col)
    $n = 0; try { $n = $Col.Count } catch { return ,@() }
    $start = 1
    try { $null = $Col.Item(0); $start = 0 } catch { $start = 1 }
    $end = if ($start -eq 0) { $n - 1 } else { $n }
    $arr = @()
    for ($i = $start; $i -le $end; $i++) {
        try { $it = $Col.Item($i); if ($null -ne $it) { $arr += $it } } catch { }
    }
    return ,$arr
}

function Get-SemanticAlias {
    param([string] $Prefix, [string] $Name, $Def)
    if ($Def -and $Def.Contains('alias') -and $Def['alias']) { return [string] $Def['alias'] }
    return "$Prefix." + ($Name -replace '[^A-Za-z0-9]', '')
}

function Get-IdByAlias {
    param($OpObject, [string] $Method, [string] $Alias)
    try { $id = $OpObject.$Method($Alias); if ($null -ne $id) { return [int] $id } } catch { }
    return -1
}

function Import-YamlFile {
    param([string] $Path)
    if (-not (Test-Path -LiteralPath $Path)) { throw "YAML file not found: $Path" }
    if (-not (Get-Module -ListAvailable -Name 'powershell-yaml')) {
        throw "The 'powershell-yaml' module is required. Install: Install-Module powershell-yaml -Scope CurrentUser"
    }
    Import-Module powershell-yaml -ErrorAction Stop
    return (Get-Content -LiteralPath $Path -Raw | ConvertFrom-Yaml)
}

# --------------------------------------------------------------------------
# Connections
# --------------------------------------------------------------------------
function Get-VaultConnections {
    (New-Object -ComObject 'MFilesAPI.MFilesClientApplication').GetVaultConnections()
}

function Show-Connections {
    $conns = ConvertTo-Array (Get-VaultConnections)
    Write-Info "Registered vault connections ($($conns.Count)):"
    foreach ($c in $conns) { '{0,-28} {1}  [{2}]' -f $c.Name, $c.ServerVaultGUID, $c.ProtocolSequence | Write-Host }
}

function Resolve-VaultGuid {
    param([string] $ConnName, [string] $Guid)
    if ($Guid) { return $Guid }
    if ($ConnName) {
        foreach ($c in (ConvertTo-Array (Get-VaultConnections))) {
            if ($c.Name -eq $ConnName) { return $c.ServerVaultGUID }
        }
    }
    return $null
}

# Find a registered vault connection by name (for its endpoint details).
function Get-RegisteredConnection {
    param([string] $ConnName)
    if (-not $ConnName) { return $null }
    foreach ($c in (ConvertTo-Array (Get-VaultConnections))) {
        if ($c.Name -eq $ConnName) { return $c }
    }
    return $null
}

# Interactive vault picker: list registered connections and let the user choose
# by number or name. Used when no -ConnectionName / connection.name is supplied,
# so the generated build command needs no editing.
function Select-Connection {
    $conns = ConvertTo-Array (Get-VaultConnections)
    if ($conns.Count -eq 0) { throw "No registered vault connections found. Register the vault in M-Files Desktop first." }
    Write-Info "Available vaults:"
    for ($i = 0; $i -lt $conns.Count; $i++) {
        '{0,3}) {1,-30} [{2}]' -f ($i + 1), $conns[$i].Name, $conns[$i].ProtocolSequence | Write-Host
    }
    while ($true) {
        $sel = (Read-Host "Which vault? (number or name)").Trim()
        if ([string]::IsNullOrWhiteSpace($sel)) { continue }
        $n = 0
        if ([int]::TryParse($sel, [ref] $n) -and $n -ge 1 -and $n -le $conns.Count) { return [string] $conns[$n - 1].Name }
        foreach ($c in $conns) { if ($c.Name -eq $sel) { return [string] $c.Name } }
        foreach ($c in $conns) { if ($c.Name.ToLowerInvariant() -eq $sel.ToLowerInvariant()) { return [string] $c.Name } }
        Write-Warn2 "  '$sel' didn't match a listed vault - enter a number or exact name."
    }
}

# Dump a vault's live metadata - object types, classes (+ their properties),
# property definitions, and value lists (+ options) - so you can see the real
# names to build an objects/tagging YAML against an EXISTING vault. Read-only.
function Show-VaultSchema {
    param($Vault)
    $dtNames = @{ 1='Text'; 2='Integer'; 3='Number'; 5='Date'; 6='Time'; 7='Timestamp'; 8='Boolean'; 9='Lookup'; 10='MultiSelectLookup'; 13='MultiLineText' }

    $objTypes = ConvertTo-Array $Vault.ObjectTypeOperations.GetObjectTypes()
    $otById = @{ 0 = 'Document' }
    foreach ($ot in $objTypes) { $otById[[int] $ot.ID] = [string] $ot.NameSingular }
    $valLists = ConvertTo-Array $Vault.ValueListOperations.GetValueLists()
    $vlById = @{}
    foreach ($vl in $valLists) { $vlById[[int] $vl.ID] = [string] $vl.NameSingular }
    $allPds = ConvertTo-Array $Vault.PropertyDefOperations.GetPropertyDefs()
    $pdById = @{}
    foreach ($pd in $allPds) { $pdById[[int] $pd.ID] = [string] $pd.Name }
    $wfById = @{}
    try { foreach ($w in (ConvertTo-Array $Vault.WorkflowOperations.GetWorkflows())) { $wfById[[int] $w.ID] = [string] $w.Name } } catch { }

    Write-Host ""
    Write-Info "=== OBJECT TYPES ==="
    foreach ($ot in ($objTypes | Sort-Object NameSingular)) {
        '  {0,-28} (id {1})' -f $ot.NameSingular, $ot.ID | Write-Host
    }

    Write-Info "=== CLASSES (each with its properties; * = required; wf: mandatory workflow) ==="
    foreach ($cl in (ConvertTo-Array $Vault.ClassOperations.GetAllObjectClasses() | Sort-Object Name)) {
        $otName = if ($otById.ContainsKey([int] $cl.ObjectType)) { $otById[[int] $cl.ObjectType] } else { "objType $($cl.ObjectType)" }
        $wfNote = ''
        try { if ([int] $cl.Workflow -gt 0) { $wfn = if ($wfById.ContainsKey([int] $cl.Workflow)) { $wfById[[int] $cl.Workflow] } else { "id $($cl.Workflow)" }; $wfNote = "  (wf: $wfn)" } } catch { }
        Write-Host ("  [{0}] {1}{2}" -f $otName, $cl.Name, $wfNote)
        $props = @()
        foreach ($apd in (ConvertTo-Array $cl.AssociatedPropertyDefs)) {
            $nm = if ($pdById.ContainsKey([int] $apd.PropertyDef)) { $pdById[[int] $apd.PropertyDef] } else { "prop $($apd.PropertyDef)" }
            $props += ($nm + $(if ($apd.Required) { '*' } else { '' }))
        }
        if ($props.Count -gt 0) { Write-Host ("        " + ($props -join ', ')) }
    }

    Write-Info "=== PROPERTIES (name : type [-> lookup target]) ==="
    foreach ($pd in ($allPds | Sort-Object Name)) {
        $t = if ($dtNames.ContainsKey([int] $pd.DataType)) { $dtNames[[int] $pd.DataType] } else { "type$($pd.DataType)" }
        $tgt = ''
        if ($pd.BasedOnValueList) {
            $vln = if ($vlById.ContainsKey([int] $pd.ValueList)) { $vlById[[int] $pd.ValueList] } else { "list $($pd.ValueList)" }
            $tgt = " -> $vln"
        }
        '  {0,-30} {1}{2}' -f $pd.Name, $t, $tgt | Write-Host
    }

    Write-Info "=== VALUE LISTS (pick options; object-type lists show existing objects) ==="
    foreach ($vl in ($valLists | Sort-Object NameSingular)) {
        $names = @()
        try { $names = @((ConvertTo-Array $Vault.ValueListItemOperations.GetValueListItems([int] $vl.ID)) | ForEach-Object { [string] $_.Name }) } catch { }
        $shown = if ($names.Count -gt 25) { (@($names[0..24]) -join ', ') + " ... (+$($names.Count - 25) more)" } else { ($names -join ', ') }
        Write-Host ("  {0} ({1}): {2}" -f $vl.NameSingular, $names.Count, $shown)
    }
    Write-Host ""
    Write-Ok "Schema dump complete - use these names in your objects/tagging YAML."
}

# Administrative server connection (Windows-integrated) - for structure. Late-bound.
function Connect-VaultAdmin {
    param([string] $Guid)
    if (-not $Guid) { throw "Admin connection needs the vault GUID (connection.vault / -VaultGuid, or a resolvable connection.name)." }
    Write-Info "Opening administrative server connection (Windows)..."
    $server = New-Object -ComObject 'MFilesAPI.MFilesServerApplication'
    $tz = New-Object -ComObject 'MFilesAPI.TimeZoneInformation'
    $tz.LoadWithCurrentTimeZone()
    [void] $server.ConnectAdministrativeEx($tz)
    $vault = $server.LogInToVaultAdministrative($Guid)
    Write-Ok "Admin connection to '$($vault.Name)' established."
    return $vault
}

# Turn a SecureString into a plain password string (freed immediately after).
function ConvertTo-PlainPassword {
    param([System.Security.SecureString] $Secret, [string] $UserName)
    if (-not $Secret) { $Secret = Read-Host -AsSecureString -Prompt "M-Files password for '$UserName'" }
    $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($Secret)
    try { return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr) }
    finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
}

# ADMINISTRATIVE connection as a named M-Files user, via a specific endpoint.
# Used for SCHEMA (structure) work - structure changes require an administrative
# login. Works against a cloud vault when the account has "Full control of vault".
function Connect-SchemaAdminUser {
    param(
        [string] $Guid, [string] $UserName, [System.Security.SecureString] $Secret,
        [string] $Prot = $Protocol, [string] $Addr = $NetAddress, [string] $End = $Port,
        [bool] $Enc = $false, [int] $Auth = 3
    )
    if (-not $Guid) { throw "Schema connection needs the vault GUID." }
    $pw = ConvertTo-PlainPassword $Secret $UserName
    Write-Info "Connecting (admin) as M-Files user '$UserName' ($Prot $Addr`:$End, encrypted=$Enc)..."
    $serverApp = New-Object -ComObject 'MFilesAPI.MFilesServerApplication'
    $tz = New-Object -ComObject 'MFilesAPI.TimeZoneInformation'; $tz.LoadWithCurrentTimeZone()
    [void] $serverApp.ConnectAdministrativeEx($tz, $Auth, $UserName, $pw, '', $null, $Prot, $Addr, $End, $Enc, '')
    $vault = $serverApp.LogInToVaultAdministrative($Guid)
    Write-Ok "Admin-connected as '$($vault.SessionInfo.AccountName)' to '$($vault.Name)'."
    return $vault
}

# NORMAL (non-administrative) connection as a named M-Files user. Late-bound COM.
# Used for OBJECT create/update. Administrative sessions stall on object
# search/create over grpc (cloud), so object work uses a normal LogInToVault -
# this is what the working pre-refactor version did.
function Connect-ObjectUser {
    param(
        [string] $Guid, [string] $UserName, [System.Security.SecureString] $Secret,
        [string] $Prot = $Protocol, [string] $Addr = $NetAddress, [string] $End = $Port,
        [bool] $Enc = $false, [int] $Auth = 3
    )
    if (-not $Guid) { throw "Object connection needs the vault GUID." }
    $pw = ConvertTo-PlainPassword $Secret $UserName
    Write-Info "Connecting as M-Files user '$UserName' ($Prot $Addr`:$End, encrypted=$Enc)..."
    $serverApp = New-Object -ComObject 'MFilesAPI.MFilesServerApplication'
    $tz = New-Object -ComObject 'MFilesAPI.TimeZoneInformation'; $tz.LoadWithCurrentTimeZone()
    # ConnectEx(TimeZone, AuthType, User, Password, Domain, SPN, Protocol, Address, Endpoint, Encrypted, LocalComputerName, AllowAnonymous)
    [void] $serverApp.ConnectEx($tz, $Auth, $UserName, $pw, '', $null, $Prot, $Addr, $End, $Enc, '', $false)
    $vault = $serverApp.LogInToVault($Guid)
    Write-Ok "Connected as '$($vault.SessionInfo.AccountName)' to '$($vault.Name)'."
    return $vault
}

# ==========================================================================
# SCHEMA CREATION (admin connection, late-bound)
# ==========================================================================
function Resolve-BasedOn {
    param($Vault, [string] $Name)
    $clean = ($Name -replace '[^A-Za-z0-9]', '')
    $otId = Get-IdByAlias $Vault.ObjectTypeOperations 'GetObjectTypeIDByAlias' "OT.$clean"
    if ($otId -ge 0) { return @{ Id = $otId; IsObjectType = $true } }
    $vlId = Get-IdByAlias $Vault.ValueListOperations 'GetValueListIDByAlias' "VL.$clean"
    if ($vlId -ge 0) { return @{ Id = $vlId; IsObjectType = $false } }
    if ($Name.ToLowerInvariant() -eq 'users') { return @{ Id = $MFBuiltInValueListUsers; IsObjectType = $false } }
    throw "based_on '$Name' not found (define its value list / object type first)."
}

function New-ValueListDef {
    param($Vault, $Def)
    $name = [string] $Def['name']; $alias = Get-SemanticAlias 'VL' $name $Def
    $id = Get-IdByAlias $Vault.ValueListOperations 'GetValueListIDByAlias' $alias
    if ($id -ge 0) {
        Write-Info "Value list '$name' exists (ID $id)."
        if (-not $DryRun) {
            try {
                $vlAdmin = $Vault.ValueListOperations.GetValueListAdmin($id)
                if (-not $vlAdmin.ObjectType.AllowAdding) {
                    $vlAdmin.ObjectType.AllowAdding = $true
                    [void] $Vault.ValueListOperations.UpdateValueListAdmin($vlAdmin)
                    Write-Ok "  enabled 'allow adding' on '$name'."
                }
            } catch { Write-Warn2 "  could not enable allow-adding on '$name': $($_.Exception.Message)" }
        }
    }
    elseif ($DryRun) { Write-Info "DryRun: would create value list '$name'."; return }
    else {
        $vla = New-Object -ComObject 'MFilesAPI.ObjTypeAdmin'
        $vla.ObjectType.NameSingular   = $name
        $vla.ObjectType.NamePlural     = $name
        $vla.ObjectType.RealObjectType = $false
        $vla.ObjectType.AllowAdding    = $true
        $vla.SemanticAliases.Value     = $alias
        $created = $Vault.ValueListOperations.AddValueListAdmin($vla)
        $id = $created.ObjectType.ID
        Write-Ok "Created value list '$name' (ID $id)."
    }

    if ($Def.Contains('items') -and $Def['items'] -and -not $DryRun) {
        $existing = @{}
        foreach ($it in (ConvertTo-Array $Vault.ValueListItemOperations.GetValueListItems($id))) {
            $existing[$it.Name.ToLowerInvariant()] = $true
        }
        foreach ($itemName in @($Def['items'])) {
            $inm = [string] $itemName
            if (-not $existing.ContainsKey($inm.ToLowerInvariant())) {
                $vli = New-Object -ComObject 'MFilesAPI.ValueListItem'
                $vli.Name = $inm; $vli.ValueListID = $id
                [void] $Vault.ValueListItemOperations.AddValueListItem($id, $vli, $false)
                Write-Ok "    + item '$inm'"
            }
        }
    }
}

function New-ObjectTypeDef {
    param($Vault, $Def)
    $name = [string] $Def['name']; $alias = Get-SemanticAlias 'OT' $name $Def
    $id = Get-IdByAlias $Vault.ObjectTypeOperations 'GetObjectTypeIDByAlias' $alias
    if ($id -ge 0) {
        Write-Info "Object type '$name' exists (ID $id)."
        if (-not $DryRun) {
            try {
                $ota = $Vault.ObjectTypeOperations.GetObjectTypeAdmin($id)
                if (-not $ota.ObjectType.AllowAdding) {
                    $ota.ObjectType.AllowAdding = $true
                    [void] $Vault.ObjectTypeOperations.UpdateObjectTypeAdmin($ota)
                    Write-Ok "  enabled 'users can create objects' on '$name'."
                }
            } catch { Write-Warn2 "  could not enable create on '$name': $($_.Exception.Message)" }
        }
        return
    }
    if ($DryRun) { Write-Info "DryRun: would create object type '$name'."; return }

    $plural = if ($Def.Contains('name_plural')) { [string] $Def['name_plural'] } else { "$name" + 's' }
    $ota = New-Object -ComObject 'MFilesAPI.ObjTypeAdmin'
    $ota.ObjectType.NameSingular   = $name
    $ota.ObjectType.NamePlural     = $plural
    $ota.ObjectType.RealObjectType = $true
    $ota.ObjectType.CanHaveFiles   = $true
    $ota.ObjectType.AllowAdding    = $true   # "Users can create objects of this type"
    $ota.SemanticAliases.Value     = $alias
    $created = $Vault.ObjectTypeOperations.AddObjectTypeAdmin($ota)
    Write-Ok "Created object type '$name' (ID $($created.ObjectType.ID))."
}

function Find-AutoLookupProp {
    param($Vault, [int] $ObjTypeId, [int] $DataType)
    foreach ($pd in (ConvertTo-Array $Vault.PropertyDefOperations.GetPropertyDefs())) {
        if ($pd.BasedOnValueList -and $pd.ValueList -eq $ObjTypeId -and $pd.DataType -eq $DataType) { return $pd.ID }
    }
    return -1
}

# Make a property searchable and usable as a view grouping level. Without these,
# a view that groups by / filters on the property fails with 0x80040486.
# Multiline text can't be a grouping level, so skip that flag for it.
function Set-PropSearchable {
    param($Pd, [int] $Dt = 0)
    try { $Pd.ObjectsSearchableByThisProperty = $true } catch { }
    if ($Dt -ne $MFDatatypeMultiLineText) { try { $Pd.AllowedAsGroupingLevel = $true } catch { } }
}

function New-PropertyDefDef {
    param($Vault, $Def)
    $name = [string] $Def['name']; $alias = Get-SemanticAlias 'PD' $name $Def
    $ops = $Vault.PropertyDefOperations
    $existingId = Get-IdByAlias $ops 'GetPropertyDefIDByAlias' $alias
    if ($existingId -ge 0) {
        if ($DryRun) { Write-Info "Property '$name' exists."; return }
        # Repair: ensure existing property is searchable / groupable.
        $adm = $ops.GetPropertyDefAdmin($existingId)
        $before = '{0}|{1}' -f $adm.PropertyDef.ObjectsSearchableByThisProperty, $adm.PropertyDef.AllowedAsGroupingLevel
        Set-PropSearchable $adm.PropertyDef $adm.PropertyDef.DataType
        $after = '{0}|{1}' -f $adm.PropertyDef.ObjectsSearchableByThisProperty, $adm.PropertyDef.AllowedAsGroupingLevel
        if ($before -ne $after) {
            try   { [void] $ops.UpdatePropertyDefAdmin($adm, $null); Write-Ok "Property '$name' exists - enabled search/grouping." }
            catch { Write-Info "Property '$name' exists (search/grouping not applicable)." }
        }
        else { Write-Info "Property '$name' exists." }
        return
    }

    $typeName = ([string] $Def['type']).ToLowerInvariant()
    if (-not $SchemaTypeMap.ContainsKey($typeName)) { throw "Unknown property type '$typeName' for '$name'." }
    $dt = $SchemaTypeMap[$typeName]
    $isLookup = ($dt -eq $MFDatatypeLookup -or $dt -eq $MFDatatypeMultiSelectLookup)

    $basedOnId = -1; $isObjType = $false
    if ($isLookup) {
        if (-not $Def.Contains('based_on')) { throw "Property '$name' is a lookup but has no 'based_on'." }
        $r = Resolve-BasedOn $Vault $Def['based_on']; $basedOnId = $r.Id; $isObjType = $r.IsObjectType
    }
    if ($DryRun) { Write-Info "DryRun: would create property '$name' ($typeName)."; return }

    if ($isLookup -and $isObjType) {
        $autoId = Find-AutoLookupProp $Vault $basedOnId $dt
        if ($autoId -ge 0) {
            $adopt = $ops.GetPropertyDefAdmin($autoId)
            $adopt.PropertyDef.Name = $name
            $adopt.SemanticAliases.Value = $alias
            Set-PropSearchable $adopt.PropertyDef $adopt.PropertyDef.DataType
            [void] $ops.UpdatePropertyDefAdmin($adopt, $null)
            Write-Ok "Adopted auto-created lookup property for '$name' (ID $autoId)."
            return
        }
    }

    $pda = New-Object -ComObject 'MFilesAPI.PropertyDefAdmin'
    $pda.PropertyDef.Name     = $name
    $pda.PropertyDef.DataType = $dt
    if ($isLookup) { $pda.PropertyDef.BasedOnValueList = $true; $pda.PropertyDef.ValueList = $basedOnId }
    Set-PropSearchable $pda.PropertyDef $dt
    $pda.SemanticAliases.Value = $alias
    $created = $ops.AddPropertyDefAdmin($pda, $null)
    Write-Ok "Created property '$name' (ID $($created.PropertyDef.ID))."
}

function Set-ClassDef {
    param($Vault, $Def)
    $name = [string] $Def['name']; $alias = Get-SemanticAlias 'CL' $name $Def
    $ops = $Vault.ClassOperations
    if ((Get-IdByAlias $ops 'GetObjectClassIDByAlias' $alias) -ge 0) { Write-Info "Class '$name' exists."; return }

    $otName = [string] $Def['object_type']
    $otId = if ($otName.ToLowerInvariant() -eq 'document') { 0 } else { (Resolve-BasedOn $Vault $otName).Id }
    if ($DryRun) { Write-Info "DryRun: would create/adopt class '$name' on '$otName'."; return }

    $defaultId = -1
    if ($otId -ne 0) {
        foreach ($ca in (ConvertTo-Array $ops.GetAllObjectClassesAdmin())) {
            if ($ca.ObjectType -eq $otId) {
                $a = ''; try { $a = [string] $ca.SemanticAliases.Value } catch { }
                if ([string]::IsNullOrWhiteSpace($a)) { $defaultId = $ca.ID; break }
            }
        }
    }
    $adopting = ($defaultId -ge 0)
    $admin = if ($adopting) { $ops.GetObjectClassAdmin($defaultId) } else { $x = New-Object -ComObject 'MFilesAPI.ObjectClassAdmin'; $x.ObjectType = $otId; $x }
    $admin.Name = $name
    $admin.SemanticAliases.Value = $alias

    foreach ($pn in @($Def['properties'])) {
        $pAlias = Get-SemanticAlias 'PD' ([string] $pn) $null
        $propId = Get-IdByAlias $Vault.PropertyDefOperations 'GetPropertyDefIDByAlias' $pAlias
        if ($propId -lt 0) { throw "Property '$pn' (alias $pAlias) not found for class '$name'." }
        $already = $false
        foreach ($apd in (ConvertTo-Array $admin.AssociatedPropertyDefs)) { if ($apd.PropertyDef -eq $propId) { $already = $true; break } }
        if (-not $already) {
            $newApd = New-Object -ComObject 'MFilesAPI.AssociatedPropertyDef'
            $newApd.PropertyDef = $propId; $newApd.Required = $false
            [void] $admin.AssociatedPropertyDefs.Add(-1, $newApd)
        }
    }

    if ($adopting) { [void] $ops.UpdateObjectClassAdmin($admin); Write-Ok "Adopted/updated class '$name' (ID $($admin.ID))." }
    else           { $created = $ops.AddObjectClassAdmin($admin); Write-Ok "Created class '$name' (ID $($created.ID))." }
}

# ==========================================================================
# WORKFLOWS (states + transitions) - admin connection, late-bound
# ==========================================================================
# States listed in YAML are referenced by name in transitions. "No State" (the
# implicit start M-Files auto-creates) is state ID 0 - a transition 'from: No
# State' is the entry point into the workflow.
function Get-StateAlias {
    param([string] $WfName, [string] $StateName)
    return 'WFS.' + ($WfName -replace '[^A-Za-z0-9]', '') + '.' + ($StateName -replace '[^A-Za-z0-9]', '')
}
function Get-TransitionAlias {
    param([string] $WfName, [string] $From, [string] $To)
    return 'WFT.' + ($WfName -replace '[^A-Za-z0-9]', '') + '.' + ($From -replace '[^A-Za-z0-9]', '') + '_' + ($To -replace '[^A-Za-z0-9]', '')
}
function Get-StateNameList {
    param($Def)
    if (-not ($Def.Contains('states') -and $Def['states'])) { return @() }
    return @($Def['states'] | ForEach-Object { if ($_ -is [System.Collections.IDictionary]) { [string] $_['name'] } else { [string] $_ } })
}
# name -> state ID, plus "No State"/blank -> 0
function Resolve-StateId {
    param($ByName, [string] $StateName)
    if ([string]::IsNullOrWhiteSpace($StateName)) { return 0 }
    $key = $StateName.ToLowerInvariant()
    if ($key -in @('no state', 'nostate', 'none', '(no state)')) { return 0 }
    if ($ByName.ContainsKey($key)) { return $ByName[$key] }
    return $null
}

function New-WorkflowDef {
    param($Vault, $Def)
    $name    = [string] $Def['name']
    $wfAlias = Get-SemanticAlias 'WF' $name $Def
    $states  = Get-StateNameList $Def
    $trans   = if ($Def.Contains('transitions') -and $Def['transitions']) { @($Def['transitions']) } else { @() }
    $wfOps   = $Vault.WorkflowOperations
    $wfId    = Get-IdByAlias $wfOps 'GetWorkflowIDByAlias' $wfAlias

    if ($DryRun) { Write-Info "DryRun: workflow '$name' ($($states.Count) state(s), $($trans.Count) transition(s))."; return }

    if ($wfId -lt 0) {
        $wfAdmin = New-Object -ComObject 'MFilesAPI.WorkflowAdmin'
        $wfAdmin.Workflow.Name = $name
        $wfAdmin.Workflow.Description = ''
        $wfAdmin.SemanticAliases.Value = $wfAlias
        foreach ($sname in $states) {
            $st = New-Object -ComObject 'MFilesAPI.StateAdmin'
            $st.Name = $sname
            $st.SemanticAliases.Value = (Get-StateAlias $name $sname)
            $wfAdmin.States.Add(-1, $st)
        }
        $created = $wfOps.AddWorkflowAdmin($wfAdmin)
        $wfId = $created.Workflow.ID
        Write-Ok "Created workflow '$name' (ID $wfId) with $($states.Count) state(s)."
    }
    else {
        Write-Info "Workflow '$name' exists (ID $wfId)."
        $wfAdmin = $wfOps.GetWorkflowAdmin($wfId)
        $added = 0
        foreach ($sname in $states) {
            $salias = Get-StateAlias $name $sname
            if ((Get-IdByAlias $wfOps 'GetWorkflowStateIDByAlias' $salias) -lt 0) {
                $st = New-Object -ComObject 'MFilesAPI.StateAdmin'
                $st.Name = $sname
                $st.SemanticAliases.Value = $salias
                $wfAdmin.States.Add(-1, $st)
                $added++
            }
        }
        if ($added -gt 0) { [void] $wfOps.UpdateWorkflowAdmin($wfAdmin); Write-Ok "  added $added new state(s) to '$name'." }
    }

    if ($trans.Count -gt 0) { Set-WorkflowTransitions $Vault $wfId $name $trans }
}

# Second pass: states now have real IDs, so add any missing transitions.
function Set-WorkflowTransitions {
    param($Vault, [int] $WfId, [string] $WfName, $Transitions)
    $wfOps   = $Vault.WorkflowOperations
    $wfAdmin = $wfOps.GetWorkflowAdmin($WfId)
    $byName  = @{}
    foreach ($s in (ConvertTo-Array $wfAdmin.States)) { $byName[$s.Name.ToLowerInvariant()] = [int] $s.ID }
    $added = 0
    foreach ($t in $Transitions) {
        $from = [string] $t['from']; $to = [string] $t['to']
        $fromId = Resolve-StateId $byName $from
        $toId   = Resolve-StateId $byName $to
        if ($null -eq $fromId) { Write-Warn2 "  transition '$from' -> '$to': unknown state '$from'."; continue }
        if ($null -eq $toId)   { Write-Warn2 "  transition '$from' -> '$to': unknown state '$to'."; continue }
        $talias = Get-TransitionAlias $WfName $from $to
        if ((Get-IdByAlias $wfOps 'GetWorkflowStateTransitionIDByAlias' $talias) -ge 0) { continue }
        $tr = New-Object -ComObject 'MFilesAPI.StateTransition'
        $tr.FromState = $fromId
        $tr.ToState   = $toId
        $tr.Name      = "$from -> $to"
        $tr.SemanticAliases.Value = $talias
        $wfAdmin.StateTransitions.Add(-1, $tr)
        $added++
    }
    if ($added -gt 0) { [void] $wfOps.UpdateWorkflowAdmin($wfAdmin); Write-Ok "  added $added transition(s) to '$WfName'." }
    else { Write-Info "  transitions up to date for '$WfName'." }
}

# ==========================================================================
# VIEWS (common/shared grouping views) - admin connection, late-bound
# ==========================================================================
# A view groups objects into a folder hierarchy by one or more properties
# (group_by). 'common: true' (default) makes it a shared view for all users.
# Views have no alias, so existence is matched by name.
function Get-ViewIdByName {
    param($Vault, [string] $Name)
    try { $views = $Vault.ViewOperations.GetViews(0, $true, 0) } catch { return -1 }
    foreach ($v in (ConvertTo-Array $views)) { if ($v.Name -eq $Name) { return [int] $v.ID } }
    return -1
}

function New-ViewDef {
    param($Vault, $Def, $PropMap)
    $name = [string] $Def['name']
    $existing = Get-ViewIdByName $Vault $name
    if ($existing -ge 0) { Write-Info "View '$name' exists (ID $existing)."; return }

    $groupBy = if ($Def.Contains('group_by') -and $Def['group_by']) { @($Def['group_by']) } else { @() }
    $common  = if ($Def.Contains('common')) { [bool] $Def['common'] } else { $true }
    if ($DryRun) { Write-Info "DryRun: $(if($common){'common'}else{'personal'}) view '$name' grouped by [$($groupBy -join ', ')]."; return }

    $view = New-Object -ComObject 'MFilesAPI.View'
    $view.Name    = $name
    $view.Common  = $common
    $view.Visible = $true
    $view.ViewType = 0
    $view.ReturnLatestVisibleVersion = $true
    foreach ($pname in $groupBy) {
        $key = ([string]$pname).ToLowerInvariant()
        $propId = if ($PropMap.ContainsKey($key)) { $PropMap[$key] } else { -1 }
        if ($propId -lt 0) { Write-Warn2 "  view '$name': unknown group-by property '$pname'."; continue }
        $ex = New-Object -ComObject 'MFilesAPI.ExpressionEx'
        $null = $ex.Expression.SetPropertyValueExpression($propId, 0, $null)
        $null = $view.Levels.Add(-1, $ex)
    }
    $created = $Vault.ViewOperations.AddView($view)
    Write-Ok "Created $(if($common){'common'}else{'personal'}) view '$name' (ID $($created.ID)) grouped by [$($groupBy -join ' > ')]."
}

function Invoke-ApplySchema {
    param($Vault, $Doc, [string] $Phase)
    Write-Info "Applying schema (phase: $Phase)$(if($DryRun){' - dry run'})..."
    $phases = if ($Phase -eq 'all') { @('valuelists','objecttypes','properties','classes','workflows','views') } else { @($Phase) }

    if ($phases -contains 'valuelists' -and $Doc.Contains('value_lists')) {
        Write-Info "-- Value lists --"
        foreach ($d in @($Doc['value_lists'])) { try { New-ValueListDef $Vault $d } catch { Write-Warn2 "value_list '$($d['name'])': $($_.Exception.Message)" } }
    }
    if ($phases -contains 'objecttypes' -and $Doc.Contains('object_types')) {
        Write-Info "-- Object types --"
        foreach ($d in @($Doc['object_types'])) { try { New-ObjectTypeDef $Vault $d } catch { Write-Warn2 "object_type '$($d['name'])': $($_.Exception.Message)" } }
    }
    if ($phases -contains 'properties' -and $Doc.Contains('property_definitions')) {
        Write-Info "-- Property definitions --"
        foreach ($d in @($Doc['property_definitions'])) { try { New-PropertyDefDef $Vault $d } catch { Write-Warn2 "property '$($d['name'])': $($_.Exception.Message)" } }
    }
    if ($phases -contains 'classes' -and $Doc.Contains('classes')) {
        Write-Info "-- Classes --"
        foreach ($d in @($Doc['classes'])) { try { Set-ClassDef $Vault $d } catch { Write-Warn2 "class '$($d['name'])': $($_.Exception.Message)" } }
    }
    if ($phases -contains 'workflows' -and $Doc.Contains('workflows')) {
        Write-Info "-- Workflows --"
        foreach ($d in @($Doc['workflows'])) { try { New-WorkflowDef $Vault $d } catch { Write-Warn2 "workflow '$($d['name'])': $($_.Exception.Message)" } }
    }
    if ($phases -contains 'views' -and $Doc.Contains('views')) {
        Write-Info "-- Views --"
        $propMap = @{}
        foreach ($p in (ConvertTo-Array $Vault.PropertyDefOperations.GetPropertyDefs())) { $propMap[$p.Name.ToLowerInvariant()] = [int] $p.ID }
        foreach ($d in @($Doc['views'])) { try { New-ViewDef $Vault $d $propMap } catch { Write-Warn2 "view '$($d['name'])': $($_.Exception.Message)" } }
    }
    Write-Ok "Schema apply complete."
}

# ==========================================================================
# OBJECTS (data) - early-bound (interop), connected as a named M-Files user
# ==========================================================================
function New-MetadataCache {
    param($Vault)
    $classesByName = @{}
    foreach ($c in (ConvertTo-Array $Vault.ClassOperations.GetAllObjectClasses())) { $classesByName[$c.Name.ToLowerInvariant()] = $c }
    $propsByName = @{}
    foreach ($p in (ConvertTo-Array $Vault.PropertyDefOperations.GetPropertyDefs())) { $propsByName[$p.Name.ToLowerInvariant()] = $p }
    return [pscustomobject]@{ Vault = $Vault; ClassesByName = $classesByName; PropsByName = $propsByName; ValueListCache = @{}; WorkflowsByName = $null; WorkflowStates = @{} }
}

# Resolve a workflow name -> its ID (built via WorkflowOperations, cached).
function Resolve-WorkflowId {
    param($Cache, [string] $Name)
    if ($null -eq $Cache.WorkflowsByName) {
        $m = @{}
        foreach ($w in (ConvertTo-Array $Cache.Vault.WorkflowOperations.GetWorkflows())) { $m[$w.Name.ToLowerInvariant()] = [int] $w.ID }
        $Cache.WorkflowsByName = $m
    }
    $k = $Name.ToLowerInvariant()
    if (-not $Cache.WorkflowsByName.ContainsKey($k)) { throw "Workflow '$Name' not found in the vault." }
    return $Cache.WorkflowsByName[$k]
}

# The StateAdmin list of a workflow (cached). Each item has .ID and .Name.
function Get-WorkflowStateList {
    param($Cache, [int] $WfId)
    if (-not $Cache.WorkflowStates.ContainsKey($WfId)) {
        $states = @()
        try { $states = ConvertTo-Array $Cache.Vault.WorkflowOperations.GetWorkflowAdmin($WfId).States } catch { }
        $Cache.WorkflowStates[$WfId] = $states
    }
    return $Cache.WorkflowStates[$WfId]
}

# Pick a workflow's initial state: the target of a transition FROM 'No State' (0),
# else the first state. Returns state ID or -1.
function Get-WorkflowInitialStateId {
    param($Cache, [int] $WfId)
    $states = Get-WorkflowStateList $Cache $WfId
    if (-not $states -or @($states).Count -eq 0) { return -1 }
    try {
        $wfAdmin = $Cache.Vault.WorkflowOperations.GetWorkflowAdmin($WfId)
        foreach ($t in (ConvertTo-Array $wfAdmin.StateTransitions)) {
            if ([int] $t.FromState -eq 0) { return [int] $t.ToState }
        }
    } catch { }
    return [int] (@($states)[0].ID)
}

# Resolve a state name within a workflow -> state ID, else -1.
function Resolve-WorkflowStateByName {
    param($Cache, [int] $WfId, [string] $Name)
    foreach ($s in (Get-WorkflowStateList $Cache $WfId)) {
        if ($s.Name.ToLowerInvariant() -eq $Name.ToLowerInvariant()) { return [int] $s.ID }
    }
    return -1
}

function Resolve-Class {
    param($Cache, [string] $Name)
    $key = $Name.ToLowerInvariant()
    if (-not $Cache.ClassesByName.ContainsKey($key)) { throw "Class '$Name' not found in the vault." }
    return $Cache.ClassesByName[$key]
}

function Resolve-PropertyDef {
    param($Cache, [string] $Name)
    $key = $Name.ToLowerInvariant()
    if (-not $Cache.PropsByName.ContainsKey($key)) { throw "Property '$Name' not found in the vault." }
    return $Cache.PropsByName[$key]
}

function Resolve-ValueListItemId {
    param($Cache, [int] $ValueListId, [string] $DisplayName)
    if (-not $Cache.ValueListCache.ContainsKey($ValueListId)) {
        $map = @{}
        foreach ($it in (ConvertTo-Array $Cache.Vault.ValueListItemOperations.GetValueListItems($ValueListId))) {
            $map[$it.Name.ToLowerInvariant()] = $it.ID
        }
        $Cache.ValueListCache[$ValueListId] = $map
    }
    $lm = $Cache.ValueListCache[$ValueListId]; $key = $DisplayName.ToLowerInvariant()
    if ($lm.ContainsKey($key)) { return $lm[$key] }
    if ($CreateMissingLookups) {
        $vli = New-Object -ComObject 'MFilesAPI.ValueListItem'; $vli.Name = $DisplayName; $vli.ValueListID = $ValueListId
        [void] $Cache.Vault.ValueListItemOperations.AddValueListItem($ValueListId, $vli, $false)
        $Cache.ValueListCache.Remove($ValueListId)
        return (Resolve-ValueListItemId $Cache $ValueListId $DisplayName)
    }
    throw "Value-list item '$DisplayName' not found in value list $ValueListId (use -CreateMissingLookups)."
}

function Resolve-LookupId {
    param($Cache, $PropDef, $Raw)
    if ($Raw -is [System.Collections.IDictionary] -and $Raw.Contains('id')) { return [int] $Raw['id'] }
    if ($Raw -is [int] -or $Raw -is [long]) { return [int] $Raw }
    return Resolve-ValueListItemId $Cache $PropDef.ValueList ([string] $Raw)
}

function Set-TypedValue {
    param($Cache, $TypedValue, $PropDef, $Raw)
    $dt = [int] $PropDef.DataType
    switch ($dt) {
        { $_ -in @(1,13) }   { $TypedValue.SetValue($dt, [string] $Raw); break }
        2                    { $TypedValue.SetValue(2, [int64] $Raw); break }
        3                    { $TypedValue.SetValue(3, [double] $Raw); break }
        8                    { $TypedValue.SetValue(8, [bool] $Raw); break }
        { $_ -in @(5,6,7) }  { $TypedValue.SetValue($dt, [datetime] $Raw); break }
        9 {
            $lk = New-Object -ComObject 'MFilesAPI.Lookup'; $lk.Item = Resolve-LookupId $Cache $PropDef $Raw; $lk.Version = -1
            $TypedValue.SetValueToLookup($lk); break
        }
        10 {
            $lks = New-Object -ComObject 'MFilesAPI.Lookups'
            foreach ($one in @($Raw)) { $l = New-Object -ComObject 'MFilesAPI.Lookup'; $l.Item = Resolve-LookupId $Cache $PropDef $one; $l.Version = -1; $null = $lks.Add(-1, $l) }
            $null = $TypedValue.SetValueToMultiSelectLookup($lks); break
        }
        default { throw "Unsupported data type ($dt) for property '$($PropDef.Name)'." }
    }
}

function Build-PropertyValues {
    param($Cache, $ObjDef, $ClassObj)
    $pvs = New-Object -ComObject 'MFilesAPI.PropertyValues'
    $pvClass = New-Object -ComObject 'MFilesAPI.PropertyValue'; $pvClass.PropertyDef = $MFBuiltInPropertyDefClass
    $lk = New-Object -ComObject 'MFilesAPI.Lookup'; $lk.Item = $ClassObj.ID; $null = $pvClass.TypedValue.SetValueToLookup($lk); $null = $pvs.Add(-1, $pvClass)
    if ($ObjDef.Contains('title') -and $ObjDef['title']) {
        $pvTitle = New-Object -ComObject 'MFilesAPI.PropertyValue'; $pvTitle.PropertyDef = $MFBuiltInPropertyDefNameOrTitle
        $null = $pvTitle.TypedValue.SetValue($MFDatatypeText, [string] $ObjDef['title']); $null = $pvs.Add(-1, $pvTitle)
    }
    # Workflow (+ initial/explicit state) - required for classes with a mandatory
    # workflow. 'workflow:' by name; optional 'state:' by name (else the workflow's
    # initial state is used).
    if ($ObjDef.Contains('workflow') -and $ObjDef['workflow']) {
        $wfId = Resolve-WorkflowId $Cache ([string] $ObjDef['workflow'])
        $pvWf = New-Object -ComObject 'MFilesAPI.PropertyValue'; $pvWf.PropertyDef = $MFBuiltInPropertyDefWorkflow
        $lkWf = New-Object -ComObject 'MFilesAPI.Lookup'; $lkWf.Item = $wfId; $null = $pvWf.TypedValue.SetValueToLookup($lkWf); $null = $pvs.Add(-1, $pvWf)
        $stId = if ($ObjDef.Contains('state') -and $ObjDef['state']) { Resolve-WorkflowStateByName $Cache $wfId ([string] $ObjDef['state']) } else { Get-WorkflowInitialStateId $Cache $wfId }
        if ($stId -ge 0) {
            $pvSt = New-Object -ComObject 'MFilesAPI.PropertyValue'; $pvSt.PropertyDef = $MFBuiltInPropertyDefState
            $lkSt = New-Object -ComObject 'MFilesAPI.Lookup'; $lkSt.Item = $stId; $null = $pvSt.TypedValue.SetValueToLookup($lkSt); $null = $pvs.Add(-1, $pvSt)
        }
    }
    if ($ObjDef.Contains('properties') -and $ObjDef['properties']) {
        foreach ($name in $ObjDef['properties'].Keys) {
            $propDef = Resolve-PropertyDef $Cache $name
            $pv = New-Object -ComObject 'MFilesAPI.PropertyValue'; $pv.PropertyDef = $propDef.ID
            $null = Set-TypedValue $Cache $pv.TypedValue $propDef $ObjDef['properties'][$name]
            $null = $pvs.Add(-1, $pv)
        }
    }
    return ,$pvs   # comma: PropertyValues is enumerable; bare 'return $pvs' unrolls it into its items
}

function Build-SourceFiles {
    param($ObjDef)
    $sf = New-Object -ComObject 'MFilesAPI.SourceObjectFiles'
    if (-not ($ObjDef.Contains('files') -and $ObjDef['files'])) { return ,$sf }
    foreach ($path in @($ObjDef['files'])) {
        if (-not (Test-Path -LiteralPath $path)) { throw "File to attach not found: $path" }
        $item = Get-Item -LiteralPath $path
        $one  = New-Object -ComObject 'MFilesAPI.SourceObjectFile'
        $one.SourceFilePath = $item.FullName; $one.Title = $item.BaseName; $one.Extension = $item.Extension.TrimStart('.')
        $null = $sf.Add(-1, $one)
    }
    return ,$sf
}

function New-VaultObject {
    param($Cache, $ObjDef, $ClassObj, $Pvs)
    $fileCount = 0
    if ($ObjDef.Contains('files') -and $ObjDef['files']) { $fileCount = @($ObjDef['files']).Count }
    $files = Build-SourceFiles $ObjDef
    $isSingleFile = ($fileCount -eq 1)
    if ($DryRun) { Write-Info "DryRun: would CREATE $($ObjDef['class']) '$($ObjDef['title'])'."; return }
    $created = $Cache.Vault.ObjectOperations.CreateNewObjectEx($ClassObj.ObjectType, $Pvs, $files, $isSingleFile, $true, $null)
    Write-Ok "Created $($ObjDef['class']) '$($ObjDef['title'])' (ID $($created.ObjVer.ObjID.ID))."
}

# Find an existing object by an explicit 'match' block, else by title within the
# class's object type. Returns its ObjID, or $null.
function Find-ExistingObject {
    param($Cache, $ObjDef, $ClassObj)
    $conds = New-Object -ComObject 'MFilesAPI.SearchConditions'
    if ($ObjDef.Contains('match') -and $ObjDef['match']) {
        foreach ($name in $ObjDef['match'].Keys) {
            $pd = Resolve-PropertyDef $Cache $name
            $cond = New-Object -ComObject 'MFilesAPI.SearchCondition'
            $expr = New-Object -ComObject 'MFilesAPI.Expression'
            $null = $expr.SetPropertyValueExpression($pd.ID, 0, $null)
            $tv = New-Object -ComObject 'MFilesAPI.TypedValue'
            $null = Set-TypedValue $Cache $tv $pd $ObjDef['match'][$name]
            $null = $cond.Set($expr, 1, $tv)
            $null = $conds.Add(-1, $cond)
        }
    }
    elseif ($ObjDef.Contains('title') -and $ObjDef['title']) {
        $cond = New-Object -ComObject 'MFilesAPI.SearchCondition'
        $expr = New-Object -ComObject 'MFilesAPI.Expression'
        $null = $expr.SetPropertyValueExpression($MFBuiltInPropertyDefNameOrTitle, 0, $null)
        $tv = New-Object -ComObject 'MFilesAPI.TypedValue'
        $null = $tv.SetValue($MFDatatypeText, [string] $ObjDef['title'])
        $null = $cond.Set($expr, 1, $tv)
        $null = $conds.Add(-1, $cond)
    }
    else { return $null }

    $results = $Cache.Vault.ObjectSearchOperations.SearchForObjectsByConditions($conds, 0, $false)
    foreach ($r in (ConvertTo-Array $results)) {
        if ($r.ObjVer.Type -eq $ClassObj.ObjectType) { return $r.ObjVer.ObjID }
    }
    return $null
}

# Version-independent comparison key for a TypedValue. Lookups are keyed by
# their item ID(s) only - the stored value carries a concrete version while the
# one we build uses Version -1, so TypedValue.CompareTo would never match.
function Get-TypedValueKey {
    param($Tv)
    if ($Tv.IsNULL()) { return '<null>' }
    $dt = [int] $Tv.DataType
    if ($dt -eq 9)  { return 'L:' + $Tv.GetValueAsLookup().Item }
    if ($dt -eq 10) {
        $ids = @(); foreach ($l in (ConvertTo-Array $Tv.GetValueAsLookups())) { $ids += [int] $l.Item }
        return 'M:' + (($ids | Sort-Object) -join ',')
    }
    return "T${dt}:" + $Tv.GetValueAsUnlocalizedText()
}

# True if the existing object already matches the desired property values (and
# files, if specified) - so we can skip the check-out/check-in cycle, which
# would otherwise bump the object version even when nothing actually changed.
function Test-ObjectMatches {
    param($Cache, $ObjID, $ObjDef, $Pvs)
    try   { $ovp = $Cache.Vault.ObjectOperations.GetLatestObjectVersionAndProperties($ObjID, $true, $true) }
    catch { return $false }   # can't read it -> treat as changed
    $have = $ovp.Properties
    foreach ($want in (ConvertTo-Array $Pvs)) {
        if ($want.PropertyDef -eq $MFBuiltInPropertyDefClass) { continue }
        $cur = $have.SearchForPropertyEx($want.PropertyDef, $true)
        if ($null -eq $cur) { if ($env:MFDEBUG) { Write-Warn2 "    diff: propDef $($want.PropertyDef) missing on object" }; return $false }
        $kw = Get-TypedValueKey $want.TypedValue; $kh = Get-TypedValueKey $cur.TypedValue
        if ($kw -ne $kh) { if ($env:MFDEBUG) { Write-Warn2 "    diff: propDef $($want.PropertyDef): want[$kw] have[$kh]" }; return $false }
    }
    if ($ObjDef.Contains('files') -and $ObjDef['files']) {
        $yaml = @($ObjDef['files'])
        $ex   = ConvertTo-Array $ovp.VersionData.Files
        if ($ex.Count -ne $yaml.Count) { if ($env:MFDEBUG) { Write-Warn2 "    diff: file count have=$($ex.Count) want=$($yaml.Count)" }; return $false }
        $exSizes = @($ex   | ForEach-Object { [int64] $_.LogicalSize } | Sort-Object)
        $waSizes = @($yaml | ForEach-Object { [int64] (Get-Item -LiteralPath $_).Length } | Sort-Object)
        for ($i = 0; $i -lt $exSizes.Count; $i++) { if ($exSizes[$i] -ne $waSizes[$i]) { if ($env:MFDEBUG) { Write-Warn2 "    diff: file size have=$($exSizes[$i]) want=$($waSizes[$i])" }; return $false } }
    }
    return $true
}

# Make a checked-out object's files match the YAML 'files:' list: remove all
# existing files, then add the specified ones. Operates on the checked-out $Ver.
function Sync-ObjectFiles {
    param($Vault, $Ver, $ObjDef)
    $yaml = @($ObjDef['files'])
    foreach ($p in $yaml) { if (-not (Test-Path -LiteralPath $p)) { throw "File to attach not found: $p" } }
    $fo     = $Vault.ObjectFileOperations
    $oo     = $Vault.ObjectOperations
    $objVer = $Ver.ObjVer
    $existing = ConvertTo-Array $Ver.Files
    # Toggle to multi-file so removing/adding isn't blocked by single-file rules.
    try { $oo.SetSingleFileObject($objVer, $false) } catch { }
    foreach ($ef in $existing) { $null = $fo.RemoveFile($objVer, $ef.FileVer) }
    foreach ($p in $yaml) {
        $item = Get-Item -LiteralPath $p
        $null = $fo.AddFile($objVer, $item.BaseName, $item.Extension.TrimStart('.'), $item.FullName)
    }
    # A single specified file → keep it a single-file document.
    if ($yaml.Count -eq 1) { try { $oo.SetSingleFileObject($objVer, $true) } catch { } }
    Write-Info "  synced $($yaml.Count) file(s) on ID $($objVer.ObjID.ID)."
}

function Update-VaultObject {
    param($Cache, $ObjID, $ObjDef, $Pvs)
    if ($DryRun) { Write-Info "DryRun: would UPDATE object ID $($ObjID.ID)."; return }
    if (Test-ObjectMatches $Cache $ObjID $ObjDef $Pvs) {
        Write-Info "Unchanged $($ObjDef['class']) '$($ObjDef['title'])' (ID $($ObjID.ID))."
        return
    }
    $oo = $Cache.Vault.ObjectOperations
    $po = $Cache.Vault.ObjectPropertyOperations
    # SetProperties cannot set/change the built-in Class property (0x80040015);
    # it is only needed at create time, so drop it before updating.
    $clsIdx = $Pvs.IndexOf($MFBuiltInPropertyDefClass)
    if ($clsIdx -ne -1) { $Pvs.Remove($clsIdx) }
    $ver = $oo.CheckOut($ObjID)
    try {
        $null = $po.SetProperties($ver.ObjVer, $Pvs)
        if ($ObjDef.Contains('files') -and $ObjDef['files']) { Sync-ObjectFiles $Cache.Vault $ver $ObjDef }
        $null = $oo.CheckIn($ver.ObjVer)
        Write-Ok "Updated $($ObjDef['class']) '$($ObjDef['title'])' (ID $($ObjID.ID))."
    } catch { try { $null = $oo.UndoCheckout($ver.ObjVer) } catch { }; throw }
}

function Invoke-Objects {
    param($Vault, $Doc)
    if (-not $Doc.Contains('objects')) { Write-Info "No 'objects:' section - nothing to do."; return }
    $cache   = New-MetadataCache $Vault
    $objects = @($Doc['objects'])
    Write-Info "Processing $($objects.Count) object(s)$(if($DryRun){' (dry run)'})..."
    $ok = 0; $fail = 0
    foreach ($objDef in $objects) {
        try {
            $classObj = Resolve-Class $cache $objDef['class']
            $pvs      = Build-PropertyValues $cache $objDef $classObj
            if ($pvs -is [Array]) { throw "Build-PropertyValues leaked $($pvs.Count) values into its output (expected a single PropertyValues)." }
            if ($objDef.Contains('id') -and $objDef['id']) {
                $objID = New-Object -ComObject 'MFilesAPI.ObjID'
                $objID.SetIDs($classObj.ObjectType, [int] $objDef['id'])
                Update-VaultObject $cache $objID $objDef $pvs
            }
            elseif (-not $AllowDuplicates) {
                $existing = Find-ExistingObject $cache $objDef $classObj
                if ($existing) { Update-VaultObject $cache $existing $objDef $pvs }
                else           { New-VaultObject   $cache $objDef $classObj $pvs }
            }
            else { New-VaultObject $cache $objDef $classObj $pvs }
            $ok++
        } catch { $fail++; Write-Warn2 "FAILED '$($objDef['title'])': $($_.Exception.Message)" }
    }
    Write-Host ""
    Write-Ok "Objects done. Success: $ok  Failed: $fail"
    if ($fail -gt 0) { $script:hadFailures = $true }
}

# ==========================================================================
# MAIN
# ==========================================================================
$script:hadFailures = $false

if ($ListConnections) { Show-Connections; return }
if (-not $YamlPath -and -not $ListSchema) { throw "YamlPath is required (or use -ListConnections / -ListSchema)." }

$doc  = if ($YamlPath) { Import-YamlFile -Path $YamlPath } else { @{} }
$conn = @{}; if ($doc.Contains('connection')) { $conn = $doc['connection'] }
$connName = if ($ConnectionName) { $ConnectionName } elseif ($conn.Contains('name'))  { $conn['name'] }  else { $null }
$vlt      = if ($VaultGuid)      { $VaultGuid }      elseif ($conn.Contains('vault')) { $conn['vault'] } else { $null }
# No vault specified anywhere -> ask (lists the registered connections to pick from),
# so the generated command doesn't need a hard-coded -ConnectionName.
if (-not $connName -and -not $vlt) { $connName = Select-Connection }
$guid = Resolve-VaultGuid -ConnName $connName -Guid $vlt
if (-not $guid) { throw "Could not resolve a vault GUID for '$connName'. Use -ListConnections to see available vaults." }

# Resolve the connection endpoint. -UseConnectionEndpoint (alias -Cloud) pulls the
# protocol/address/port/encryption from the registered connection, so the tool can
# reach a cloud vault instead of the localhost default.
$prot = $Protocol; $addr = $NetAddress; $end = $Port; $enc = [bool] $Encrypted; $auth = $AuthType
$rc = $null
if ($UseConnectionEndpoint) {
    $rc = Get-RegisteredConnection $connName
    if (-not $rc) { throw "-UseConnectionEndpoint/-Cloud needs a registered connection.name / -ConnectionName; '$connName' not found (see -ListConnections)." }
    $prot = $rc.ProtocolSequence; $addr = $rc.NetworkAddress; $end = $rc.Endpoint
    if (-not $Encrypted) { $enc = [bool] $rc.EncryptedConnection }
    Write-Info "Using registered endpoint of '$connName': $prot $addr`:$end (encrypted=$enc)."
}

# Prompt for the M-Files user when we'll connect as one and none was given, so
# different runs can use different accounts. Default to the connection's stored
# user when it has one.
$needsUser = $SingleLogin -or $UseConnectionEndpoint -or $ListSchema -or (-not $SchemaOnly)
if ($needsUser -and -not $User) {
    $default = if ($rc -and $rc.UserName) { [string] $rc.UserName } else { '' }
    $label   = if ($default) { "M-Files user [$default]" } else { "M-Files user" }
    $entered = Read-Host $label
    $User    = if ([string]::IsNullOrWhiteSpace($entered)) { $default } else { $entered.Trim() }
    if (-not $User) { throw "A username is required (enter one at the prompt, or pass -User)." }
}

# -ListSchema: connect read-only (normal login - no vault-admin needed) and dump the
# vault's object types, classes, properties, and value lists, then exit.
if ($ListSchema) {
    $svault = Connect-ObjectUser -Guid $guid -UserName $User -Secret $Password -Prot $prot -Addr $addr -End $end -Enc $enc -Auth $auth
    Show-VaultSchema $svault
    return
}

if ($SingleLogin -or $UseConnectionEndpoint) {
    # One administrative connection as the M-Files user, used for both phases
    # (matches the known-working build). The account needs "Full control of vault".
    $vault = Connect-SchemaAdminUser -Guid $guid -UserName $User -Secret $Password -Prot $prot -Addr $addr -End $end -Enc $enc -Auth $auth
    if ($ApplySchema)     { Invoke-ApplySchema $vault $doc $SchemaPhase }
    if (-not $SchemaOnly) { Invoke-Objects     $vault $doc }
}
else {
    if ($ApplySchema) {
        $adminVault = Connect-VaultAdmin -Guid $guid
        Invoke-ApplySchema $adminVault $doc $SchemaPhase
    }
    if (-not $SchemaOnly) {
        $objVault = Connect-ObjectUser -Guid $guid -UserName $User -Secret $Password
        Invoke-Objects $objVault $doc
    }
}

if ($script:hadFailures) { exit 1 }
