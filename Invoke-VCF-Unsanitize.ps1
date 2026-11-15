<#
.SYNOPSIS
  Restore original names, IPs and identifiers in a delivered report (Excel, CSV, HTML, text) using the token key created by Invoke-VcfSanitize.ps1 (or vcfr_sanitize.py).
  Run this AT THE CUSTOMER SITE with the original KEY.csv. Also looks up a token or name.
.DESCRIPTION
  - Reads KEY.csv (Token,Type,Original,FirstSeenIn). Every token such as HOST-0007 or VM-00042 (any letter case) is replaced by its original value.
  - Technical values preserved by the Rev2 sanitizer are passed through unchanged because they were never tokenized.
  - This script remains compatible with older key files and can restore older sanitized reports; values previously replaced with TEXT-REMOVED cannot be recovered.
  - Host and vCenter names that were split into LABEL + DOMAIN tokens (HOST-0007.DOM-01) are restored to the full original name.
  - Excel sheet names are restored too. Cell formatting added by the report is kept where the ImportExcel module can preserve it; otherwise values are restored on a fresh sheet.
  - Requires: PowerShell 7+ (or Windows PowerShell 5.1); ImportExcel module for .xlsx (Install-Module ImportExcel -Scope CurrentUser).
.PARAMETER Path     Files (.xlsx, .csv, .html, .md, .txt, .json) or folders to restore.
.PARAMETER KeyFile  The KEY.csv used when the data was sanitized.
.PARAMETER OutDir   Output folder (files are named <name>_RESTORED.<ext>).
.PARAMETER Lookup   Instead of restoring files, print what one or more tokens or names map to (e.g. -Lookup HOST-0007,VC-01,vdcpa).
.EXAMPLE
  .\Invoke-VCF-Unsanitize-Rev2.ps1 -Path .\Customer_VCF_9.1_Readiness_Data.xlsx,.\Customer_VCF_9.1_Readiness_Report.html -KeyFile .\KEY.csv -OutDir .\final
.EXAMPLE
  .\Invoke-VCF-Unsanitize-Rev2.ps1 -KeyFile .\KEY.csv -Lookup HOST-0007,VC-01
.NOTES
  Written by Presidio. Not executed by the author in the customer's environment: test on a copy first.
#>
[CmdletBinding(DefaultParameterSetName = 'Restore')]
param(
  [Parameter(ParameterSetName = 'Restore', Mandatory)][string[]]$Path,
  [Parameter(Mandatory)][string]$KeyFile,
  [Parameter(ParameterSetName = 'Restore', Mandatory)][string]$OutDir,
  [Parameter(ParameterSetName = 'Lookup', Mandatory)][string[]]$Lookup
)
$ErrorActionPreference = 'Stop'
if (-not (Test-Path $KeyFile)) { throw "Key file not found: $KeyFile" }
$Rev = @{}                                    # UPPER token -> original
foreach ($r in (Import-Csv $KeyFile)) { $Rev[$r.Token.ToUpper()] = $r.Original }
if ($Rev.Count -eq 0) { throw 'Key file is empty.' }
$Types = @{}; foreach ($r in (Import-Csv $KeyFile)) { $Types[$r.Token.ToUpper()] = $r.Type }

# ---------------- lookup mode ----------------
if ($PSCmdlet.ParameterSetName -eq 'Lookup') {
  foreach ($q in $Lookup) {
    if ($Rev.ContainsKey($q.ToUpper())) { '{0,-14} {1,-10} {2}' -f $q.ToUpper(), $Types[$q.ToUpper()], $Rev[$q.ToUpper()] }
    else {
      $hits = $Rev.Keys | Where-Object { $Rev[$_] -like "*$q*" } | Select-Object -First 25
      if (-not $hits) { "${q}: not found" } else { foreach ($h in $hits) { '{0,-14} {1,-10} {2}' -f $h, $Types[$h], $Rev[$h] } }
    }
  }
  return
}

# ---------------- restore ----------------
$prefixes = ($Rev.Keys | ForEach-Object { $_.Split('-')[0] } | Select-Object -Unique) -join '|'
$reTok = [regex]("(?i)(?<![A-Za-z0-9])(?:$prefixes)-\d{2,6}(?!\d)")
$ev = [System.Text.RegularExpressions.MatchEvaluator] { param($m) $k = $m.Value.ToUpper(); if ($Rev.ContainsKey($k)) { $Rev[$k] } else { $m.Value } }
function Restore-Text($v) { if ($null -eq $v) { return $v }; $s = [string]$v; if ($s.Length -lt 4 -or -not $s.Contains('-')) { return $v }; return $reTok.Replace($s, $ev) }

function Restore-Rows($rows) {
  if (-not $rows -or @($rows).Count -eq 0) { return @() }
  $cols = @($rows[0].PSObject.Properties.Name)
  $maps = @{}
  foreach ($c in $cols) { $m = @{}; foreach ($v in ($rows | ForEach-Object { $_.$c } | Where-Object { $_ -is [string] -and $_ } | Select-Object -Unique)) { $m[$v] = Restore-Text $v }; $maps[$c] = $m }
  foreach ($r in $rows) {
    $o = [ordered]@{}
    foreach ($c in $cols) { $nc = Restore-Text $c; $v = $r.$c; $o[$nc] = if ($v -is [string] -and $v) { $maps[$c][$v] } else { $v } }
    [pscustomobject]$o
  }
}

New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$files = foreach ($p in $Path) { if (Test-Path $p -PathType Container) { Get-ChildItem $p -File | Where-Object { $_.Extension -in '.xlsx','.xlsm','.csv','.txt','.html','.htm','.md','.json','.log' } } else { Get-Item $p } }
foreach ($f in $files) {
  $dst = Join-Path $OutDir ("{0}_RESTORED{1}" -f ($f.BaseName -replace '_(SANITIZED|RESTORED)$',''), $f.Extension)
  switch ($f.Extension.ToLower()) {
    { $_ -in '.xlsx','.xlsm' } {
      if (-not (Get-Module -ListAvailable ImportExcel)) { throw 'ImportExcel module is required for .xlsx (Install-Module ImportExcel -Scope CurrentUser)' }
      Import-Module ImportExcel
      if (Test-Path $dst) { Remove-Item $dst -Force }
      foreach ($sh in (Get-ExcelSheetInfo -Path $f.FullName).Name) {
        $out = Restore-Rows (Import-Excel -Path $f.FullName -WorksheetName $sh)
        $name = (Restore-Text $sh); $name = ($name -replace '[\\/\?\*\[\]:]', '_'); $name = $name.Substring(0, [math]::Min(31, $name.Length))
        if (@($out).Count -gt 0) { $out | Export-Excel -Path $dst -WorksheetName $name } }
    }
    '.csv' { (Restore-Rows (Import-Csv $f.FullName)) | Export-Csv -Path $dst -NoTypeInformation -Encoding UTF8 }
    default { Set-Content -Path $dst -Value (Restore-Text (Get-Content $f.FullName -Raw)) -Encoding UTF8 }
  }
  $left = 0
  if ($f.Extension -in '.csv','.txt','.html','.htm','.md','.json') { $left = $reTok.Matches((Get-Content $dst -Raw)).Count }
  Write-Host ("  restored {0} -> {1}{2}" -f $f.Name, (Split-Path $dst -Leaf), $(if ($left) { "   ($left token(s) not in this key were left as-is)" } else { '' }))
}
Write-Host "Done. Restored files contain real customer names - handle and store them accordingly." -ForegroundColor Yellow
