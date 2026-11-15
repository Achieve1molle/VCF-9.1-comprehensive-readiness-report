<#
.SYNOPSIS
  Sanitize RVTools workbooks and PowerCLI / OEM CSV files: replaces host names, VM names, vCenters, clusters, networks, datastores, IPs, MACs, UUIDs
  and serial numbers with consistent tokens (HOST-0001, VM-00042, VC-01, IP-00007 ...). Run this AT THE CUSTOMER SITE before sending data.
.DESCRIPTION
  - The same original always gets the same token, across every file and every run that reuses the same key file.
  - KEY.csv (Token,Type,Original,FirstSeenIn) NEVER leaves the customer site. It is the same format used by the Python tool vcfr_sanitize.py, so keys are interchangeable.
  - Technical compatibility fields are preserved verbatim. General free-text columns are replaced with TEXT-REMOVED; Description is preserved only in pcli_* and bmc_* evidence files. The vLicense / vMetaData sheets and backup / ID columns are dropped.
  - Requires: PowerShell 7+ (recommended) or Windows PowerShell 5.1, and the ImportExcel module for .xlsx (Install-Module ImportExcel -Scope CurrentUser). CSV files need nothing extra.
.PARAMETER Path       One or more files (.xlsx, .csv, .txt, .html, .md, .json) or folders (all supported files inside).
.PARAMETER KeyFile    Token key file. Created if missing, extended if it exists. KEEP IT AT THE CUSTOMER SITE.
.PARAMETER OutDir     Output folder for the sanitized copies (files are named <name>_SANITIZED.<ext>).
.PARAMETER Verify     After sanitizing, scan the outputs for any original value still present and report LEAK / CLEAN.
.EXAMPLE
  .\Invoke-VCF-Sanitize-Rev2.ps1 -Path .\RVTools_export.xlsx,.\vcfr-collect -KeyFile .\KEY.csv -OutDir .\to-send -Verify
.NOTES
  Written by Presidio. Not executed by the author in the customer's environment: test on a copy first and review the output before sending.
#>
[CmdletBinding()]
param(
  [Parameter(Mandatory)][string[]]$Path,
  [Parameter(Mandatory)][string]$KeyFile,
  [Parameter(Mandatory)][string]$OutDir,
  [switch]$Verify
)
$ErrorActionPreference = 'Stop'

# ---------------- configuration (mirrors vcfr_sanitize.py) ----------------
$Prefix = @{ host=@('HOST',4); vm=@('VM',5); vcenter=@('VC',2); cluster=@('CL',3); datacenter=@('DC',2); network=@('NET',4); datastore=@('DS',3); folder=@('FLD',4);
             switch=@('SW',3); ip=@('IP',5); mac=@('MAC',5); uuid=@('UUID',5); serial=@('SER',4); domain=@('DOM',2); email=@('MAIL',3) }
$ColType = @{
 'vm'='vm'; 'dns name'='host'; 'host'='host'; 'vi sdk server'='vcenter'; 'vcenter'='vcenter'; 'vc'='vcenter'; 'vc_fqdn'='vcenter'; 'server'='vcenter';
 'cluster'='cluster'; 'cluster name'='cluster'; 'datacenter'='datacenter'; 'folder'='folder'; 'resource pool'='folder'; 'vapp'='folder';
 'network'='network'; 'network #1'='network'; 'network #2'='network'; 'network #3'='network'; 'network #4'='network'; 'network #5'='network'; 'network #6'='network'; 'network #7'='network'; 'network #8'='network';
 'port group'='network'; 'switch'='switch'; 'datastore'='datastore'; 'service tag'='serial'; 'serial number'='serial'; 'servicetag'='serial'; 'serial'='serial'; 'serialnumber'='serial'; 'device service tag'='serial';
 'devicename'='host'; 'device name'='host'; 'servername'='host'; 'hostname'='host'; 'vm uuid'='uuid'; 'vi sdk uuid'='uuid'; 'smbios uuid'='uuid'; 'uuid'='uuid'; 'object id'='uuid'; 'vmref'='uuid';
 'mac address'='mac'; 'mac'='mac'; 'primary ip address'='ip'; 'ip address'='ip'; 'ipv4 address'='ip'; 'ip'='ip'; 'domain'='domain'; 'bmcaddress'='ip' }
$Redact   = @('annotation','notes','description','custom fields','customization info','contact','admin name','cost unit','labels')
$DropCols = @('key','password','rubrik_lastbackup','created_at','cpi_metadata_version','compiling','director','deployment','instance_group','job','id','index','xdconfig')
$SkipSheets = @('vmetadata','vlicense')
$GenericDirs = @('fcd','vsan','vdisks','vmkdump','iso','templates','templ','isos','scratch','.sdd.sf','lost+found','tmp','dvsdata')
$NoIpCol  = 'version|build|bios|firmware|driver|tools|api|product|release'
$Generic  = @('datastore','staging','default','production','storage','network','cluster','server','windows','linux','template','templates','management','resources','resource','standard','virtual','machine','guest','testing','general','vm network','vmkernel','management network','vmotion','vsan','vm network 2')
# Compatibility evidence is intentionally preserved verbatim. Environment identity remains tokenized.
$PreserveTech = @(
 'model','manufacturer','vendor','cpu model','cpumodel','cpu series','cpuseries','processor','processor type',
 'guest os','guestos','os','os name','osname','os version','osversion','guest full name','guestfullname',
 'bios version','biosversion','firmware','firmware version','firmwareversion','driver','driver version','driverversion',
 'vid','did','svid','ssid','class','protocol','media','media type','mediatype','transport','device class',
 'base image','baseimage','addon','add-on','addon version','addonversion','add-on version','hardware support','hardware_support',
 'encryption mode','encryptionmode','require secure boot','requiresecureboot','require exec installed only','requireexecinstalledonly',
 'tpm present','tpmpresent','drtm enabled','drtmenabled','tpm version','tpmversion','lockdown mode','lockdownmode','ssh running','sshrunning',
 'ntp servers','ntpservers','ntp running','ntprunning','ntp policy','ntppolicy','boot mode','bootmode','secure boot','secureboot',
 'power state','powerstate','health','device firmware','devicefirmware','bmc model','bmcmodel','bmc firmware','bmcfirmware','dimm types','dimmtypes'
)
$TechnicalPrefixes = @('pcli_','bmc_')
function Test-PreserveTechnical([string]$sheet, [string]$col) {
  $sh = $sheet.Trim().ToLower(); $lc = (($col.Trim().ToLower() -replace '_',' ') -replace '\s+',' ')
  $normalized = $PreserveTech | ForEach-Object { (($_ -replace '_',' ') -replace '\s+',' ') }
  if ($normalized -contains $lc) { return $true }
  if (($TechnicalPrefixes | Where-Object { $sh.StartsWith($_) }).Count -gt 0 -and $lc -in 'description','device') { return $true }
  return $false
}

$reIP   = [regex]'(?<![\d.])(?:\d{1,3}\.){3}\d{1,3}(?![\d.]*\d)'
$reMAC  = [regex]'\b[0-9A-Fa-f]{2}(?:[:-][0-9A-Fa-f]{2}){5}\b'
$reUUID = [regex]'\b[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}\b'
$reMAIL = [regex]'\b[\w.+-]+@[\w-]+(?:\.[\w-]+)+\b'
$reWORD = [regex]'[A-Za-z0-9_][A-Za-z0-9_.\-]*'
$reTOKEN = [regex]('(?<![A-Za-z0-9])(?:' + (($Prefix.Values | ForEach-Object { $_[0] }) -join '|') + ')-\d{2,6}(?!\d)')
$rePATH = [regex]'\[([^\]\[]+)\]\s*([^/\\\s\]\[]+)((?:/[^\s]*)?)'

# ---------------- key store ----------------
$script:Fwd = @{}          # "type|lower" -> token
$script:Rev = @{}          # token -> @(type, original)
$script:Src = @{}
$script:Count = @{}; foreach ($k in $Prefix.Keys) { $script:Count[$k] = 0 }
$script:Multi = @{}        # lower multi-word original -> token
$script:MultiRx = $null
$script:DomRx = $null; $script:DomN = -1

function Test-MultiWord([string]$o, [string]$t) { return ($o -match '[\s/\\\[\]()]') -and ($t -notin 'ip','mac','uuid','email') }

function Load-Key([string]$f) {
  if (-not (Test-Path $f)) { return }
  foreach ($r in (Import-Csv $f)) {
    $script:Fwd["$($r.Type)|$($r.Original.ToLower())"] = $r.Token; $script:Rev[$r.Token] = @($r.Type, $r.Original); $script:Src[$r.Token] = $r.FirstSeenIn
    $n = [int]($r.Token.Split('-')[-1]); if ($n -gt $script:Count[$r.Type]) { $script:Count[$r.Type] = $n }
    if (Test-MultiWord $r.Original $r.Type) { $script:Multi[$r.Original.ToLower()] = $r.Token }
  }
  $script:MultiRx = $null
}
function Save-Key([string]$f) {
  $rows = foreach ($tok in ($script:Rev.Keys | Sort-Object { $script:Rev[$_][0] }, { [int]($_.Split('-')[-1]) })) {
    [pscustomobject]@{ Token=$tok; Type=$script:Rev[$tok][0]; Original=$script:Rev[$tok][1]; FirstSeenIn=$script:Src[$tok] } }
  $rows | Export-Csv -Path $f -NoTypeInformation -Encoding UTF8
}
function Get-Token([string]$typ, [string]$orig, [string]$where = '') {
  $o = $orig.Trim(); if (-not $o) { return $orig }
  $k = "$typ|$($o.ToLower())"
  if ($script:Fwd.ContainsKey($k)) { return $script:Fwd[$k] }
  $script:Count[$typ]++; $pre = $Prefix[$typ]
  $tok = '{0}-{1}' -f $pre[0], ([string]$script:Count[$typ]).PadLeft($pre[1], '0')
  $script:Fwd[$k] = $tok; $script:Rev[$tok] = @($typ, $o); $script:Src[$tok] = $where
  if (Test-MultiWord $o $typ) { $script:Multi[$o.ToLower()] = $tok; $script:MultiRx = $null }
  return $tok
}
function Get-Entity([string]$typ, $value, [string]$where) {
  if ($null -eq $value) { return $value }
  $v = ([string]$value).Trim(); if (-not $v -or $v -ieq 'nan') { return $value }
  if ($reIP.IsMatch($v) -and $reIP.Match($v).Value -eq $v) { return Get-Token 'ip' $v $where }
  if ($typ -in 'host','vcenter' -and $v.Contains('.')) {
    $i = $v.IndexOf('.'); $lab = $v.Substring(0, $i); $dom = $v.Substring($i + 1)
    return (Get-Token $typ $lab $where) + '.' + (Get-Token 'domain' $dom $where)
  }
  return Get-Token $typ $v $where
}

function Lookup-Word([string]$w) {
  if ($reTOKEN.IsMatch($w) -and $reTOKEN.Match($w).Value -eq $w) { return $w }
  $tail = ''; while ($w.EndsWith('.') -or $w.EndsWith('-')) { $tail = $w.Substring($w.Length - 1) + $tail; $w = $w.Substring(0, $w.Length - 1) }
  $lw = $w.ToLower()
  foreach ($t in 'host','vm','vcenter','cluster','datacenter','network','datastore','folder','switch','serial','domain') { $k = "$t|$lw"; if ($script:Fwd.ContainsKey($k)) { return $script:Fwd[$k] + $tail } }
  foreach ($sep in '.','_','-') {                                   # name + suffix, e.g. VMNAME.vmx, VMNAME_2.vmdk
    $cut = $w.LastIndexOf($sep)
    while ($cut -gt 3) {
      $head = $w.Substring(0, $cut).ToLower()
      if ($head.Length -ge 5) { foreach ($t in 'host','vm','vcenter','cluster','network','datastore','folder','switch','serial') { $k = "$t|$head"; if ($script:Fwd.ContainsKey($k)) { return $script:Fwd[$k] + $w.Substring($cut) + $tail } } }
      if ($cut -le 0) { break }; $cut = $w.LastIndexOf($sep, $cut - 1)
    }
  }
  if ($w.Contains('.')) {
    $i = $w.IndexOf('.'); $lab = $w.Substring(0, $i).ToLower(); $dom = $w.Substring($i + 1).ToLower()
    foreach ($t in 'host','vm','vcenter') { $a = $script:Fwd["$t|$lab"]; $d = $script:Fwd["domain|$dom"]; if ($a -and $d) { return "$a.$d$tail" } }
  }
  return $w + $tail
}
function Replace-Path([string]$text) {
  return $rePATH.Replace($text, {
    param($m)
    $ds = $m.Groups[1].Value; $folder = $m.Groups[2].Value; $rest = $m.Groups[3].Value
    $dtok = if ($reTOKEN.IsMatch($ds) -and $reTOKEN.Match($ds).Value -eq $ds) { $ds } else { Get-Token 'datastore' $ds }
    if (-not $folder -or ($reTOKEN.IsMatch($folder) -and $reTOKEN.Match($folder).Value -eq $folder) -or $GenericDirs -contains $folder.ToLower() -or $folder.Length -lt 5) { return "[$dtok] $folder$rest" }
    $ftok = Get-Entity 'vm' $folder ''
    $parts = $rest.Split('/') | ForEach-Object { if ($_ -ieq $folder) { $ftok } else { [regex]::Replace($_, '^' + [regex]::Escape($folder) + '(?=[._-])', $ftok, 'IgnoreCase') } }
    return "[$dtok] $ftok" + ($parts -join '/')
  })
}
function Scrub-Text($text, [string]$col = '') {
  if ($null -eq $text) { return $text }
  $s = [string]$text; if (-not $s -or $s -ieq 'nan') { return $text }
  if ($s.Contains('[')) { $s = Replace-Path $s }
  if ($script:Multi.Count -gt 0) {
    if ($null -eq $script:MultiRx) { $names = $script:Multi.Keys | Sort-Object Length -Descending | ForEach-Object { [regex]::Escape($_) }; $script:MultiRx = [regex]('(?i)(?<![A-Za-z0-9_])(?:' + ($names -join '|') + ')(?![A-Za-z0-9_])') }
    $s = $script:MultiRx.Replace($s, { param($m) $script:Multi[$m.Value.ToLower()] })
  }
  $s = $reMAIL.Replace($s, { param($m) Get-Token 'email' $m.Value })
  $s = $reUUID.Replace($s, { param($m) Get-Token 'uuid' $m.Value })
  $s = $reMAC.Replace($s,  { param($m) Get-Token 'mac' $m.Value })
  if ($col -notmatch $NoIpCol) { $s = $reIP.Replace($s, { param($m) Get-Token 'ip' $m.Value }) }
  $s = $reWORD.Replace($s, { param($m) Lookup-Word $m.Value })
  $nd = @($script:Fwd.Keys | Where-Object { $_ -like 'domain|*' }).Count
  if ($nd -gt 0) {
    if ($null -eq $script:DomRx -or $script:DomN -ne $nd) {
      $ds = $script:Fwd.Keys | Where-Object { $_ -like 'domain|*' } | ForEach-Object { $_.Substring(7) } | Sort-Object Length -Descending | ForEach-Object { [regex]::Escape($_) }
      $script:DomRx = [regex]('(?i)(?<=[A-Za-z0-9\]])\.(' + ($ds -join '|') + ')(?![A-Za-z0-9-])'); $script:DomN = $nd }
    $s = $script:DomRx.Replace($s, { param($m) '.' + $script:Fwd["domain|$($m.Groups[1].Value.ToLower())"] })
  }
  return $s
}

# ---------------- table handling ----------------
function Get-ColType([string]$sheet, [string]$col) {
  $lc = $col.Trim().ToLower(); $sh = $sheet.Trim().ToLower()
  if ($sh -eq 'vdatastore' -and $lc -eq 'name') { return 'datastore' }
  if ($sh -eq 'vcluster' -and $lc -eq 'name') { return 'cluster' }
  if ($sh -eq 'dvswitch' -and $lc -eq 'name') { return 'switch' }
  if ($ColType.ContainsKey($lc)) { return $ColType[$lc] }
  return $null
}
function Convert-Table($rows, [string]$sheet, [string]$where) {
  if (-not $rows -or @($rows).Count -eq 0) { return @() }
  $cols = @($rows[0].PSObject.Properties.Name) | Where-Object { $DropCols -notcontains $_.Trim().ToLower() }
  # pass 1: register typed names so free-text scrub can recognise them
  foreach ($c in $cols) { $t = Get-ColType $sheet $c; if ($t -and $t -notin 'ip','mac' -and -not (Test-PreserveTechnical $sheet $c) -and $Redact -notcontains $c.Trim().ToLower()) {
      foreach ($v in ($rows | ForEach-Object { $_.$c } | Where-Object { $_ } | Select-Object -Unique)) { [void](Get-Entity $t $v $where) } } }
  # pass 2: transform each distinct value once per column
  $maps = @{}
  foreach ($c in $cols) {
    $lc = $c.Trim().ToLower(); $t = Get-ColType $sheet $c; $m = @{}
    foreach ($v in ($rows | ForEach-Object { $_.$c } | Where-Object { $null -ne $_ -and "$_" -ne '' } | Select-Object -Unique)) {
      $key = [string]$v
      if (Test-PreserveTechnical $sheet $c) { $m[$key] = $v }
      elseif ($Redact -contains $lc) { $m[$key] = 'TEXT-REMOVED' }
      elseif ($t) { $m[$key] = Get-Entity $t $v $where }
      elseif ($v -is [string]) { $m[$key] = Scrub-Text $v $c }
      else { $m[$key] = $v }
    }
    $maps[$c] = $m
  }
  foreach ($r in $rows) {
    $o = [ordered]@{}
    foreach ($c in $cols) { $v = $r.$c; $o[$c] = if ($null -eq $v -or "$v" -eq '') { $v } else { $maps[$c][[string]$v] } }
    [pscustomobject]$o
  }
}

# ---------------- main ----------------
Load-Key $KeyFile
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$files = foreach ($p in $Path) { if (Test-Path $p -PathType Container) { Get-ChildItem $p -File | Where-Object { $_.Extension -in '.xlsx','.xlsm','.csv','.txt','.html','.htm','.md','.json','.log' } } else { Get-Item $p } }
$outs = @()
foreach ($f in $files) {
  $where = $f.Name; $dst = Join-Path $OutDir ("{0}_SANITIZED{1}" -f ($f.BaseName -replace '_(SANITIZED|RESTORED)$',''), $f.Extension)
  switch ($f.Extension.ToLower()) {
    { $_ -in '.xlsx','.xlsm' } {
      if (-not (Get-Module -ListAvailable ImportExcel)) { throw 'ImportExcel module is required for .xlsx (Install-Module ImportExcel -Scope CurrentUser)' }
      Import-Module ImportExcel
      $names = (Get-ExcelSheetInfo -Path $f.FullName).Name | Where-Object { $SkipSheets -notcontains $_.Trim().ToLower() }
      if (Test-Path $dst) { Remove-Item $dst -Force }
      foreach ($sh in $names) {
        $rows = Import-Excel -Path $f.FullName -WorksheetName $sh
        $out = Convert-Table $rows $sh $where
        if (@($out).Count -gt 0) { $out | Export-Excel -Path $dst -WorksheetName $sh.Substring(0, [math]::Min(31, $sh.Length)) -AutoSize:$false } }
    }
    '.csv' { $out = Convert-Table (Import-Csv $f.FullName) ([IO.Path]::GetFileNameWithoutExtension($f.Name)) $where; $out | Export-Csv -Path $dst -NoTypeInformation -Encoding UTF8 }
    default { $t = Get-Content $f.FullName -Raw; Set-Content -Path $dst -Value (Scrub-Text $t) -Encoding UTF8 }
  }
  $outs += $dst; Write-Host ("  sanitized {0} -> {1}" -f $f.Name, (Split-Path $dst -Leaf))
}
Save-Key $KeyFile
Write-Host ("key file: {0} ({1:N0} tokens)  KEEP THIS AT THE CUSTOMER SITE" -f $KeyFile, $script:Rev.Count) -ForegroundColor Yellow

# ---------------- verify ----------------
if ($Verify) {
  $single = New-Object 'System.Collections.Generic.HashSet[string]'; $multi = @()
  foreach ($tok in $script:Rev.Keys) { $t, $o = $script:Rev[$tok]; if ($o.Length -lt 5 -or $t -in 'domain','datacenter','switch' -or $Generic -contains $o.ToLower()) { continue }
    if ($o -match '\s') { $multi += $o.ToLower() } else { [void]$single.Add($o.ToLower()) } }
  $origIps = @($script:Rev.Values | Where-Object { $_[0] -eq 'ip' } | ForEach-Object { $_[1] })
  $bad = 0
  foreach ($o in $outs) {
    $chunks = @()
    if ($o -like '*.xlsx') { Import-Module ImportExcel; foreach ($sh in (Get-ExcelSheetInfo -Path $o).Name) { $chunks += $sh; foreach ($r in (Import-Excel -Path $o -WorksheetName $sh)) { foreach ($p in $r.PSObject.Properties) { if ($null -ne $p.Value) { $chunks += [string]$p.Value } } } } }
    elseif ($o -like '*.csv') { foreach ($r in (Import-Csv $o)) { foreach ($p in $r.PSObject.Properties) { if ($p.Value) { $chunks += [string]$p.Value } } } }
    else { $chunks += (Get-Content $o -Raw) }
    $found = New-Object 'System.Collections.Generic.HashSet[string]'; $ips = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($c in ($chunks | Select-Object -Unique)) {
      $lc = $c.ToLower()
      foreach ($w in $reWORD.Matches($lc)) { $x = $w.Value.TrimEnd('.', '-'); if ($single.Contains($x)) { [void]$found.Add($x) } else { foreach ($sep in '.','_','-') { $i = $x.LastIndexOf($sep); if ($i -gt 0 -and $single.Contains($x.Substring(0, $i))) { [void]$found.Add($x.Substring(0, $i)); break } } } }
      foreach ($n in $multi) { if ($lc.Contains($n)) { [void]$found.Add($n) } }
      foreach ($m in $reIP.Matches($c)) { if ($origIps -contains $m.Value) { [void]$ips.Add($m.Value) } }
    }
    $state = if ($found.Count -eq 0 -and $ips.Count -eq 0) { 'CLEAN' } else { 'LEAK' }
    Write-Host ("{0}: {1} original names, {2} original IPs still present - {3}" -f (Split-Path $o -Leaf), $found.Count, $ips.Count, $state) -ForegroundColor $(if ($state -eq 'CLEAN') { 'Green' } else { 'Red' })
    ($found | Select-Object -First 10) + ($ips | Select-Object -First 10) | ForEach-Object { Write-Host "     $_" }
    $bad += $found.Count + $ips.Count
  }
  if ($bad -gt 0) { Write-Warning 'Original values remain. Do NOT send these files until the leak is resolved.' }
}
