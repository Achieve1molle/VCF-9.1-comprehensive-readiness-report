#requires -Version 7.2
<#+
.SYNOPSIS
  Unified read-only PowerCLI evidence collector for VCF 9.1 readiness.
.DESCRIPTION
  Work item 1 framework combining the former readiness, host deep-dive, and
  active-memory collectors. The framework preserves the established pcli_*.csv
  contract and centralizes profiles, dataset registration, connection handling,
  logging, append writers, checkpoints, and collection orchestration.

  vCenters are read from -VCenterList. All credentials are acquired and all
  connections are preflighted before data collection begins. Work items 3-7 will harden resume, logging, NIC handling, missing
  evidence, and representative-hardware inference.
.EXAMPLE
  .\Collect-VCF-Evidence.ps1 -VCenterList .\vcenters.txt -SharedCredential -CollectionProfile Full -IncludePci -OutDir C:\VCF-Evidence
.EXAMPLE
  .\Collect-VCF-Evidence.ps1 -VCenterList .\vcenters.txt -CollectionProfile MemoryOnly -Days 30 -OutDir C:\VCF-Evidence
.EXAMPLE
  .\Collect-VCF-Evidence.ps1 -VCenterList .\vcenters.txt -CollectionProfile Standard -DryRun
.NOTES
  Requires VMware.PowerCLI 13.x for live collection. Uses read-only Get-* and
  read-only ESXCLI/CIS calls. No configuration changes are performed.
#>
[CmdletBinding(SupportsShouldProcess=$false)]
param(
  [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$VCenterList,
  [switch]$SharedCredential,
  [string]$CredentialMap,
  [Alias('PasswordFile')][string]$ClearTextCredentialFile='',
  [switch]$AllowClearTextCredentials,
  [string]$DefaultUserName='',
  [ValidateSet('Stop','Skip')][string]$ConnectionFailureAction='Skip',
  [ValidateSet('Standard','Full','MemoryOnly','VmFeaturesOnly','Resume')][string]$CollectionProfile='Standard',
  [string]$OutDir='.\vcfr-collect',
  [pscredential]$Credential,
  [switch]$IncludePci,
  [switch]$SkipVmFeatures,
  [ValidateRange(1,5000)][int]$VmFeatureSampleLimit=50,
  [ValidateRange(1,365)][int]$Days=30,
  [ValidateRange(0,1000000)][int]$MaxVMs=0,
  [switch]$SkipVmStats,
  [switch]$DryRun,
  [string]$ResumeRunId='',
  [switch]$ForceRecollect,
  [ValidateRange(1,10000)][int]$FlushRowCount=100,
  [ValidateRange(1,5000)][int]$VmBatchSize=100,
  [ValidateRange(5,300)][int]$HeartbeatSeconds=30,
  [ValidateRange(1,10)][int]$MaxAttempts=3,
  [ValidateRange(1,300)][int]$InitialRetrySeconds=2,
  [ValidateRange(1,3600)][int]$OperationTimeoutSeconds=600,
  [ValidateRange(60,7200)][int]$HostTimeoutSeconds=900,
  [ValidateRange(0,1440)][int]$VCenterTimeoutMinutes=0,
  [ValidateRange(300,14400)][int]$StallTimeoutSeconds=2700,
  [ValidateRange(1,10)][int]$MaxOperationRetries=2,
  [ValidateRange(1,250)][int]$HostBatchSize=25,
  [switch]$ResumeLatest,
  [switch]$ForceResume,
  [switch]$DisableEvidenceZip,
  [string]$EvidenceZipPath='',
  [switch]$IncludeCollectorStateInZip,
  [switch]$NoConsoleHardening,
  [switch]$NoConsoleOutput,
  [switch]$RunInCurrentProcess,
  [Parameter(DontShow)][switch]$InternalCollectorMode,
  [ValidateRange(0,10)][int]$MaxAutomaticRestarts=3,
  [ValidateRange(5,300)][int]$SupervisorPollSeconds=15
)

# Unified single-file entry point. By default this script supervises a hidden,
# noninteractive child instance of itself. Use -RunInCurrentProcess only for
# interactive troubleshooting.
if(-not $InternalCollectorMode -and -not $RunInCurrentProcess){
  if($Credential){throw 'A live PSCredential cannot be serialized to the supervised child process. Use -PasswordFile for unattended mode or add -RunInCurrentProcess for interactive troubleshooting.'}
  New-Item -ItemType Directory -Force -Path $OutDir|Out-Null
  $self=$PSCommandPath
  if(-not$self){throw 'Unable to resolve the unified collector script path.'}
  $supervisorLog=Join-Path $OutDir 'collector-supervisor.log'
  $stdout=Join-Path $OutDir 'collector-console.out.log'
  $stderr=Join-Path $OutDir 'collector-console.err.log'
  function Add-ChildValue([Collections.Generic.List[string]]$List,[string]$Name,$Value){if($null-ne$Value-and[string]$Value-ne''){$List.Add("-$Name");$List.Add([string]$Value)}}
  function Add-ChildSwitch([Collections.Generic.List[string]]$List,[string]$Name,[bool]$Value){if($Value){$List.Add("-$Name")}}
  for($attempt=0;$attempt-le$MaxAutomaticRestarts;$attempt++){
    $profile=if($attempt-eq 0){$CollectionProfile}else{'Resume'}
    $child=[Collections.Generic.List[string]]::new()
    foreach($x in @('-NoLogo','-NoProfile','-NonInteractive','-File',$self,'-InternalCollectorMode')){$child.Add($x)}
    Add-ChildValue $child 'VCenterList' $VCenterList
    Add-ChildValue $child 'CredentialMap' $CredentialMap
    Add-ChildValue $child 'ClearTextCredentialFile' $ClearTextCredentialFile
    Add-ChildValue $child 'DefaultUserName' $DefaultUserName
    Add-ChildValue $child 'ConnectionFailureAction' $ConnectionFailureAction
    Add-ChildValue $child 'CollectionProfile' $profile
    Add-ChildValue $child 'OutDir' $OutDir
    Add-ChildValue $child 'VmFeatureSampleLimit' $VmFeatureSampleLimit
    Add-ChildValue $child 'Days' $Days
    Add-ChildValue $child 'MaxVMs' $MaxVMs
    Add-ChildValue $child 'ResumeRunId' $ResumeRunId
    Add-ChildValue $child 'FlushRowCount' $FlushRowCount
    Add-ChildValue $child 'VmBatchSize' $VmBatchSize
    Add-ChildValue $child 'HeartbeatSeconds' $HeartbeatSeconds
    Add-ChildValue $child 'MaxAttempts' $MaxAttempts
    Add-ChildValue $child 'InitialRetrySeconds' $InitialRetrySeconds
    Add-ChildValue $child 'OperationTimeoutSeconds' $OperationTimeoutSeconds
    Add-ChildValue $child 'HostTimeoutSeconds' $HostTimeoutSeconds
    Add-ChildValue $child 'VCenterTimeoutMinutes' $VCenterTimeoutMinutes
    Add-ChildValue $child 'StallTimeoutSeconds' $StallTimeoutSeconds
    Add-ChildValue $child 'MaxOperationRetries' $MaxOperationRetries
    Add-ChildValue $child 'HostBatchSize' $HostBatchSize
    Add-ChildValue $child 'EvidenceZipPath' $EvidenceZipPath
    foreach($sw in @(
      @('SharedCredential',[bool]$SharedCredential),@('AllowClearTextCredentials',[bool]$AllowClearTextCredentials),@('IncludePci',[bool]$IncludePci),
      @('SkipVmFeatures',[bool]$SkipVmFeatures),@('SkipVmStats',[bool]$SkipVmStats),@('DryRun',[bool]$DryRun),@('ForceRecollect',[bool]$ForceRecollect),
      @('ResumeLatest',[bool]$ResumeLatest),@('ForceResume',[bool]$ForceResume),@('DisableEvidenceZip',[bool]$DisableEvidenceZip),
      @('IncludeCollectorStateInZip',[bool]$IncludeCollectorStateInZip),@('NoConsoleHardening',[bool]$NoConsoleHardening),@('NoConsoleOutput',[bool]$NoConsoleOutput))){Add-ChildSwitch $child $sw[0] $sw[1]}
    Add-Content -LiteralPath $supervisorLog -Value "$(Get-Date -Format o) Attempt=$($attempt+1) Profile=$profile Starting hidden noninteractive child."
    $p=Start-Process pwsh -ArgumentList $child.ToArray() -PassThru -WindowStyle Hidden -RedirectStandardOutput $stdout -RedirectStandardError $stderr
    $lastActivity=Get-Date
    while(-not$p.HasExited){
      Start-Sleep -Seconds $SupervisorPollSeconds;$p.Refresh()
      $files=@($stdout,$stderr,(Join-Path $OutDir 'collector.log'),(Join-Path $OutDir 'collector-events.jsonl'),(Join-Path $OutDir 'checkpoint-state.json'))|Where-Object{Test-Path -LiteralPath $_}
      if($files){$latest=($files|Get-Item|Sort-Object LastWriteTimeUtc -Descending|Select-Object -First 1).LastWriteTimeUtc;if($latest-gt$lastActivity.ToUniversalTime()){$lastActivity=$latest.ToLocalTime()}}
      if(((Get-Date)-$lastActivity).TotalSeconds-ge$StallTimeoutSeconds){
        Add-Content -LiteralPath $supervisorLog -Value "$(Get-Date -Format o) No output or checkpoint activity for $StallTimeoutSeconds seconds. Stopping PID $($p.Id) for checkpoint resume."
        Stop-Process -Id $p.Id -Force;break
      }
    }
    $p.WaitForExit();$exitCode=$p.ExitCode
    Add-Content -LiteralPath $supervisorLog -Value "$(Get-Date -Format o) Attempt=$($attempt+1) ChildExitCode=$exitCode"
    if($exitCode-eq 0-or$exitCode-eq 2){exit $exitCode}
    if($attempt-eq$MaxAutomaticRestarts){throw "Collector did not complete after $($attempt+1) attempts. Review $supervisorLog, $stdout, and $stderr."}
    Start-Sleep -Seconds 10
  }
  exit 5
}

Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
$script:OriginalInformationPreference=$InformationPreference
$InformationPreference='Continue'
$PSDefaultParameterValues['*:Confirm']=$false
$PSDefaultParameterValues['*:ErrorAction']='Stop'
$env:POWERSHELL_TELEMETRY_OPTOUT='1'
function Assert-NonInteractiveCollector {
  $path=$MyInvocation.PSCommandPath
  if(-not $path){$path=$PSCommandPath}
  if(-not $path -or -not(Test-Path -LiteralPath $path)){return}
  $forbidden='(?im)^\s*(Read-Host|Pause|\$Host\.UI\.(Prompt|PromptForChoice))\b'
  $content=Get-Content -LiteralPath $path -Raw
  $matches=[regex]::Matches($content,$forbidden)
  if($matches.Count){throw "Interactive command detected in collector execution path: $($matches[0].Value.Trim())"}
}
Assert-NonInteractiveCollector
function Disable-ConsoleSelectionPause {
  if($NoConsoleHardening -or -not $IsWindows){return}
  try{
    if(-not('VcfrConsoleNative' -as [type])){
      Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class VcfrConsoleNative {
  [DllImport("kernel32.dll", SetLastError=true)] public static extern IntPtr GetStdHandle(int nStdHandle);
  [DllImport("kernel32.dll", SetLastError=true)] public static extern bool GetConsoleMode(IntPtr hConsoleHandle, out uint lpMode);
  [DllImport("kernel32.dll", SetLastError=true)] public static extern bool SetConsoleMode(IntPtr hConsoleHandle, uint dwMode);
}
'@
    }
    $h=[VcfrConsoleNative]::GetStdHandle(-10);[uint32]$mode=0
    if([VcfrConsoleNative]::GetConsoleMode($h,[ref]$mode)){
      $ENABLE_EXTENDED_FLAGS=0x0080;$ENABLE_QUICK_EDIT_MODE=0x0040
      $newMode=($mode-bor$ENABLE_EXTENDED_FLAGS)-band(-bnot$ENABLE_QUICK_EDIT_MODE)
      [void][VcfrConsoleNative]::SetConsoleMode($h,[uint32]$newMode)
    }
  }catch{ }
}
Disable-ConsoleSelectionPause
$script:CollectorVersion='4.1.2-final-transport-hotfix'
if($ResumeRunId){
  if($ResumeRunId -notmatch '^[0-9a-fA-F-]{36}$'){throw 'ResumeRunId must be a GUID from a prior manifest or checkpoint'}
  $script:RunId=$ResumeRunId.ToLowerInvariant()
}else{$script:RunId=[guid]::NewGuid().ToString()}
$script:RunStartedUtc=[datetime]::UtcNow
$script:OutDir=[IO.Path]::GetFullPath($OutDir)
$script:LogPath=Join-Path $script:OutDir 'collector.log'
$script:EventLogPath=Join-Path $script:OutDir 'collector-events.jsonl'
$script:SummaryCsvPath=Join-Path $script:OutDir 'collector-run-summary.csv'
$script:SummaryJsonPath=Join-Path $script:OutDir 'collector-run-summary.json'
$script:DatasetStatusPath=Join-Path $script:OutDir 'dataset-status.json'
$script:Metrics=[ordered]@{Info=0;OperationalWarnings=0;SecurityWarnings=0;EvidenceGaps=0;UnavailableObjects=0;QueryFailures=0;Errors=0;Retries=0;Timeouts=0;Heartbeats=0;ProgressEvents=0;RowsCommitted=0;BatchesCommitted=0;Operations=0}
$script:LastCheckpointUtc=$null
$script:OriginalProgressPreference=$ProgressPreference
$script:OriginalConfirmPreference=$ConfirmPreference
$ProgressPreference='SilentlyContinue'
$ConfirmPreference='None' 
$script:CheckpointPath=Join-Path $script:OutDir 'checkpoint-state.json'
$script:ManifestPath=Join-Path $script:OutDir ("run-manifest-{0}.json" -f $script:RunId)
$script:DatasetStatus=[ordered]@{}
$script:RunErrors=[Collections.Generic.List[string]]::new()
$script:Buffers=@{}
$script:PendingKeys=@{}
$script:PersistedKeys=@{}
$script:CompletedObjects=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$script:StateDir=Join-Path $script:OutDir '.collector-state'
$script:RunStateDir=Join-Path $script:StateDir $script:RunId
$script:JournalPath=Join-Path $script:RunStateDir 'checkpoint-journal.jsonl'
$script:StateSnapshotPath=Join-Path $script:RunStateDir 'state.json'
$script:BatchDir=Join-Path $script:RunStateDir 'batches'
$script:InvocationFingerprint='' 
$script:Credentials=@{}
$script:PreflightServers=@{}
$script:PreflightStatus=[Collections.Generic.List[object]]::new()
$script:DatasetRegistry=[ordered]@{
  pcli_vcenter       = @{ File='pcli_vcenter.csv';       Profiles=@('Standard','Full','Resume') }
  pcli_clusters      = @{ File='pcli_clusters.csv';      Profiles=@('Standard','Full','Resume') }
  pcli_hosts         = @{ File='pcli_hosts.csv';         Profiles=@('Standard','Full','Resume') }
  pcli_vmk           = @{ File='pcli_vmk.csv';           Profiles=@('Standard','Full','Resume') }
  pcli_vds           = @{ File='pcli_vds.csv';           Profiles=@('Standard','Full','Resume') }
  pcli_datastores    = @{ File='pcli_datastores.csv';    Profiles=@('Standard','Full','Resume') }
  pcli_extensions    = @{ File='pcli_extensions.csv';    Profiles=@('Standard','Full','Resume') }
  pcli_vsan          = @{ File='pcli_vsan.csv';          Profiles=@('Standard','Full','Resume') }
  pcli_pci           = @{ File='pcli_pci.csv';           Profiles=@('Full','Resume') }
  pcli_vm_encryption = @{ File='pcli_vm_encryption.csv'; Profiles=@('Standard','Full','VmFeaturesOnly','Resume') }
  pcli_vm_features   = @{ File='pcli_vm_features.csv';   Profiles=@('Standard','Full','VmFeaturesOnly','Resume') }
  pcli_vm_coverage   = @{ File='pcli_vm_coverage.csv';   Profiles=@('Standard','Full','VmFeaturesOnly','Resume') }
  pcli_boot          = @{ File='pcli_boot.csv';          Profiles=@('Full','Resume') }
  pcli_security      = @{ File='pcli_security.csv';      Profiles=@('Full','Resume') }
  pcli_vibs          = @{ File='pcli_vibs.csv';          Profiles=@('Full','Resume') }
  pcli_nicfw         = @{ File='pcli_nicfw.csv';         Profiles=@('Full','Resume') }
  pcli_dns           = @{ File='pcli_dns.csv';           Profiles=@('Full','Resume') }
  pcli_ntp           = @{ File='pcli_ntp.csv';           Profiles=@('Full','Resume') }
  pcli_images        = @{ File='pcli_images.csv';        Profiles=@('Full','Resume') }
  pcli_hsm           = @{ File='pcli_hsm.csv';           Profiles=@('Full','Resume') }
  pcli_vsandisks     = @{ File='pcli_vsandisks.csv';     Profiles=@('Full','Resume') }
  pcli_mem_vm        = @{ File='pcli_mem_vm.csv';        Profiles=@('Full','MemoryOnly','Resume') }
  pcli_mem_host      = @{ File='pcli_mem_host.csv';      Profiles=@('Full','MemoryOnly','Resume') }
  pcli_nvme          = @{ File='pcli_nvme.csv';          Profiles=@('Full','MemoryOnly','Resume') }
  pcli_vm_shared_disks = @{ File='pcli_vm_shared_disks.csv'; Profiles=@('Standard','Full','VmFeaturesOnly','Resume') }
  pcli_key_providers   = @{ File='pcli_key_providers.csv';   Profiles=@('Full','Resume') }
}
$script:DatasetContracts=[ordered]@{
  pcli_images=@('vCenter','Cluster','LifecycleMode','DesiredImageVersion','ComplianceStatus','VendorAddon','Components','LastComplianceCheckUtc','HardwareSupportManager','CollectionStatus','CollectionError')
  pcli_extensions=@('vCenter','Name','Version','Company','Type','Status','CollectionStatus','CollectionError')
  pcli_vsan=@('vCenter','Cluster','Enabled','Architecture','Stretched','Witness','FaultDomains','Health','CapacityGiB','FreeGiB','CollectionStatus','CollectionError')
  pcli_vsandisks=@('vCenter','Cluster','Host','DiskGroup','Device','Model','Firmware','CapacityGiB','Role','Health','CollectionStatus','CollectionError')
  pcli_nvme=@('vCenter','Cluster','Host','Device','Model','Serial','Namespace','Firmware','CapacityGiB','MediaType','UsageRole','VID','DID','SVID','SSID','CollectionStatus','CollectionError')
  pcli_hsm=@('vCenter','Cluster','Provider','Version','ComplianceStatus','LastCheckUtc','CollectionStatus','CollectionError')
  pcli_vm_features=@('vCenter','VM','Uuid','Firmware','EfiSecureBoot','VBS','IOMMU','vTPM','PciPassthrough','PowerState','CollectionStatus','CollectionError')
  pcli_vm_encryption=@('vCenter','VM','Uuid','Encryption','VmConfigEncrypted','EncryptedDisks','KeyProviderId','PowerState','CollectionStatus','CollectionError')
  pcli_vm_shared_disks=@('vCenter','VM','Uuid','DiskLabel','ControllerBusSharing','SharingMode','BackingFile','CollectionStatus','CollectionError')
  pcli_pci=@('vCenter','Host','Cluster','Device','Vendor','VID','DID','SVID','SSID','Driver','Class')
  pcli_security=@('vCenter','Cluster','Host','EncryptionMode','RequireSecureBoot','SecureBoot','TpmPresent','TpmVersion','TpmMode','Attestation','LockdownMode','SshRunning','CertificateStatus','CollectionStatus','CollectionError')
}
$script:DatasetTelemetry=[ordered]@{}
function Set-DatasetStatus {
  param([string]$Dataset,[ValidateSet('Produced','SupportedEmpty','UnsupportedApi','PermissionDenied','ConnectionFailure','FilteredOut','NotImplemented','QueryFailed')][string]$Status,[string]$Error='',[string]$VCenter='',[int]$Rows=0,[double]$DurationSeconds=0)
  if(-not $script:DatasetTelemetry.Contains($Dataset)){$script:DatasetTelemetry[$Dataset]=[Collections.Generic.List[object]]::new()}
  $entry=[pscustomobject]@{Dataset=$Dataset;VCenter=$VCenter;Status=$Status;Rows=$Rows;DurationSeconds=[math]::Round($DurationSeconds,3);Error=$Error;UpdatedUtc=[datetime]::UtcNow.ToString('o')}
  $existing=@($script:DatasetTelemetry[$Dataset]);$same=@($existing|Where-Object{$_.VCenter-eq$VCenter})
  if($Status-eq'Produced'){
    foreach($item in $same){[void]$script:DatasetTelemetry[$Dataset].Remove($item)}
  }elseif(@($same|Where-Object Status -eq 'Produced').Count){return}
  $script:DatasetTelemetry[$Dataset].Add($entry)
}
function Test-DatasetContract {
  param([string]$Dataset,[object]$Row)
  if(-not $script:DatasetContracts.Contains($Dataset)){return}
  $missing=@($script:DatasetContracts[$Dataset]|Where-Object{$Row.PSObject.Properties.Name -notcontains $_})
  if($missing.Count){throw "Dataset $Dataset row is missing contract columns: $($missing -join ', ')"}
}


function Write-CollectorEvent {
  param(
    [ValidateSet('DEBUG','INFO','WARN','ERROR','CHECKPOINT','PROGRESS')][string]$Level,
    [string]$Message,[string]$EventType='Message',[string]$Target='',
    [string]$VCenterName='',[string]$HostName='',[string]$Dataset='',[string]$Operation='',
    [double]$ElapsedSeconds=0,[int]$Rows=0,[int]$Attempt=0,[string]$ExceptionType=''
  )
  switch($Level){
    'INFO'{$script:Metrics.Info++}
    'WARN'{
      if($EventType -eq 'CredentialFileLoaded'){$script:Metrics.SecurityWarnings++}
      elseif($EventType -in @('EvidenceGap','DatasetGap')){$script:Metrics.EvidenceGaps++}
      elseif($EventType -eq 'UnavailableObject'){$script:Metrics.UnavailableObjects++}
      elseif($EventType -eq 'QueryFailed'){$script:Metrics.QueryFailures++}
      else{$script:Metrics.OperationalWarnings++}
    }
    'ERROR'{$script:Metrics.Errors++}
    'PROGRESS'{if($EventType -eq 'Heartbeat'){$script:Metrics.Heartbeats++}else{$script:Metrics.ProgressEvents++}}
  }
  $event=[ordered]@{
    TimestampUtc=[datetime]::UtcNow.ToString('o');RunId=$script:RunId;Level=$Level;EventType=$EventType
    Target=$Target;VCenter=$VCenterName;Host=$HostName;Dataset=$Dataset;Operation=$Operation
    ElapsedSeconds=[math]::Round($ElapsedSeconds,3);Rows=$Rows;Attempt=$Attempt
    ExceptionType=$ExceptionType;Message=$Message
  }
  $line='{0} [{1}] RunId={2} Event={3} Target={4} Operation={5} {6}' -f $event.TimestampUtc,$Level,$script:RunId,$EventType,$Target,$Operation,$Message
  if(-not $NoConsoleOutput){
    $color=switch($Level){'ERROR'{'Red'}'WARN'{'Yellow'}'PROGRESS'{'Cyan'}'CHECKPOINT'{'Green'}default{'Gray'}}
    Write-Host $line -ForegroundColor $color
  }
  if(-not $DryRun){
    Add-Content -LiteralPath $script:LogPath -Value $line -Encoding utf8
    [IO.File]::AppendAllText($script:EventLogPath,(($event|ConvertTo-Json -Compress -Depth 8)+[Environment]::NewLine),[Text.UTF8Encoding]::new($false))
  }
}
function Write-CollectorLog {
  param([ValidateSet('DEBUG','INFO','WARN','ERROR','CHECKPOINT','PROGRESS')][string]$Level,[string]$Message,[string]$Target='')
  Write-CollectorEvent -Level $Level -Message $Message -Target $Target
}
function Invoke-WithRetry {
  param(
    [Parameter(Mandatory)][scriptblock]$ScriptBlock,[Parameter(Mandatory)][string]$Operation,
    [string]$Target='',[string]$VCenterName='',[string]$HostName='',[string]$Dataset='',
    [int]$Attempts=$MaxAttempts,[int]$InitialDelaySeconds=$InitialRetrySeconds
  )
  $script:Metrics.Operations++
  $started=[datetime]::UtcNow;$delay=$InitialDelaySeconds
  for($attempt=1;$attempt-le$Attempts;$attempt++){
    try{
      Write-CollectorEvent INFO 'Operation started' OperationStarted $Target $VCenterName $HostName $Dataset $Operation 0 0 $attempt
      $result=& $ScriptBlock
      $elapsed=([datetime]::UtcNow-$started).TotalSeconds
      if($elapsed-gt$OperationTimeoutSeconds){$script:Metrics.Timeouts++;throw [TimeoutException]::new("Operation exceeded advisory timeout of $OperationTimeoutSeconds seconds (elapsed $([math]::Round($elapsed,1)))")}
      Write-CollectorEvent INFO 'Operation completed' OperationCompleted $Target $VCenterName $HostName $Dataset $Operation $elapsed 0 $attempt
      return $result
    }catch{
      $elapsed=([datetime]::UtcNow-$started).TotalSeconds;$type=$_.Exception.GetType().FullName
      if($_.Exception -is [TimeoutException]){$script:Metrics.Timeouts++}
      if($attempt-ge$Attempts){Write-CollectorEvent ERROR $_.Exception.Message OperationFailed $Target $VCenterName $HostName $Dataset $Operation $elapsed 0 $attempt $type;throw}
      $script:Metrics.Retries++;Write-CollectorEvent WARN "Retrying after $delay second(s): $($_.Exception.Message)" OperationRetry $Target $VCenterName $HostName $Dataset $Operation $elapsed 0 $attempt $type
      Start-Sleep -Seconds $delay;$delay=[math]::Min(30,$delay*2)
    }
  }
}
function Invoke-WithHeartbeat {
  param([Parameter(Mandatory)][scriptblock]$ScriptBlock,[Parameter(Mandatory)][string]$Operation,[string]$Target='',[string]$VCenterName='',[string]$HostName='',[string]$Dataset='')
  # A timer event writes heartbeats independently while the PowerCLI call remains
  # in the current runspace, preserving modules, VIServer sessions, and variables.
  $started=[datetime]::UtcNow
  $timer=[Timers.Timer]::new([math]::Max(5000,$HeartbeatSeconds*1000))
  $timer.AutoReset=$true
  $messageData=[pscustomobject]@{RunId=$script:RunId;EventLogPath=$script:EventLogPath;LogPath=$script:LogPath;Operation=$Operation;Target=$Target;VCenter=$VCenterName;Host=$HostName;Dataset=$Dataset;StartedUtc=$started;LastCheckpoint=$script:LastCheckpointUtc}
  $sourceId="VCFHeartbeat-$($script:RunId)-$([guid]::NewGuid().ToString('N'))"
  $subscription=Register-ObjectEvent -InputObject $timer -EventName Elapsed -SourceIdentifier $sourceId -MessageData $messageData -Action {
    $d=$event.MessageData;$elapsed=([datetime]::UtcNow-$d.StartedUtc).TotalSeconds
    $record=[ordered]@{TimestampUtc=[datetime]::UtcNow.ToString('o');RunId=$d.RunId;Level='PROGRESS';EventType='Heartbeat';Target=$d.Target;VCenter=$d.VCenter;Host=$d.Host;Dataset=$d.Dataset;Operation=$d.Operation;ElapsedSeconds=[math]::Round($elapsed,3);Rows=0;Attempt=0;ExceptionType='';Message="Operation running; last checkpoint=$($d.LastCheckpoint)"}
    [IO.File]::AppendAllText($d.EventLogPath,(($record|ConvertTo-Json -Compress -Depth 6)+[Environment]::NewLine),[Text.UTF8Encoding]::new($false))
    [IO.File]::AppendAllText($d.LogPath,("{0} [PROGRESS] RunId={1} Event=Heartbeat Target={2} Operation={3} {4}" -f $record.TimestampUtc,$d.RunId,$d.Target,$d.Operation,$record.Message)+[Environment]::NewLine,[Text.UTF8Encoding]::new($false))
  }
  try{
    $timer.Start()
    $result=& $ScriptBlock
    $elapsed=([datetime]::UtcNow-$started).TotalSeconds
    if($elapsed-gt$OperationTimeoutSeconds){$script:Metrics.Timeouts++;throw [TimeoutException]::new("Operation exceeded advisory timeout of $OperationTimeoutSeconds seconds (elapsed $([math]::Round($elapsed,1)))")}
    $result
  }finally{
    $timer.Stop();Unregister-Event -SourceIdentifier $sourceId -ErrorAction SilentlyContinue
    Remove-Job -Id $subscription.Id -Force -ErrorAction SilentlyContinue;$timer.Dispose()
  }
}
function Write-RunSummary {
  param([string]$Status)
  $summary=[ordered]@{RunId=$script:RunId;CollectorVersion=$script:CollectorVersion;Status=$Status;StartedUtc=$script:RunStartedUtc.ToString('o');EndedUtc=[datetime]::UtcNow.ToString('o');ElapsedSeconds=[math]::Round(([datetime]::UtcNow-$script:RunStartedUtc).TotalSeconds,2);VCentersRequested=@($script:VCenterNames).Count;VCentersReady=@($script:PreflightStatus|Where-Object Status -eq 'Ready').Count
    Info=$script:Metrics.Info;OperationalWarnings=$script:Metrics.OperationalWarnings;SecurityWarnings=$script:Metrics.SecurityWarnings
    EvidenceGaps=$script:Metrics.EvidenceGaps;UnavailableObjects=$script:Metrics.UnavailableObjects;QueryFailures=$script:Metrics.QueryFailures
    Errors=$script:Metrics.Errors;Retries=$script:Metrics.Retries;Timeouts=$script:Metrics.Timeouts;Heartbeats=$script:Metrics.Heartbeats
    ProgressEvents=$script:Metrics.ProgressEvents;Operations=$script:Metrics.Operations;RowsCommitted=$script:Metrics.RowsCommitted
    BatchesCommitted=$script:Metrics.BatchesCommitted;DatasetGaps=@($script:DatasetStatus.Values|Where-Object{$_.Status-ne'Produced'}).Count}
  Write-AtomicText $script:SummaryJsonPath ($summary|ConvertTo-Json -Depth 6)
  [pscustomobject]$summary|Export-Csv -LiteralPath $script:SummaryCsvPath -NoTypeInformation -Encoding utf8
}
function Get-ActiveDatasetNames {
  $profile=if($CollectionProfile -eq 'Resume'){'Resume'}else{$CollectionProfile}
  @($script:DatasetRegistry.Keys | Where-Object { $script:DatasetRegistry[$_].Profiles -contains $profile })
}
function Get-Sha256Text {
  param([string]$Text)
  $sha=[Security.Cryptography.SHA256]::Create()
  try{return ([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($Text))).Replace('-','').ToLowerInvariant())}finally{$sha.Dispose()}
}
function Get-EvidenceKey {
  param([string]$Dataset,[object]$Row,[string]$SourceVCenter,[string]$SourceHostKey)
  $parts=[ordered]@{Dataset=$Dataset;SourceVCenter=$SourceVCenter;SourceHostKey=$SourceHostKey}
  foreach($prop in $Row.PSObject.Properties|Sort-Object Name){$parts[$prop.Name]=$prop.Value}
  Get-Sha256Text ($parts|ConvertTo-Json -Compress -Depth 8)
}
function Write-AtomicText {
  param([string]$Path,[string]$Content)
  $dir=Split-Path -Parent $Path;if(-not(Test-Path $dir)){New-Item -ItemType Directory -Force $dir|Out-Null}
  $tmp="$Path.$([guid]::NewGuid().ToString('N')).tmp"
  [IO.File]::WriteAllText($tmp,$Content,[Text.UTF8Encoding]::new($false))
  Move-Item -LiteralPath $tmp -Destination $Path -Force
}
function Add-CheckpointEvent {
  param([string]$Event,[string]$Scope,[string]$Name,[hashtable]$Data=@{})
  $record=[ordered]@{TimestampUtc=[datetime]::UtcNow.ToString('o');RunId=$script:RunId;Event=$Event;Scope=$Scope;Name=$Name;Data=$Data}
  $line=$record|ConvertTo-Json -Compress -Depth 8
  [IO.File]::AppendAllText($script:JournalPath,$line+[Environment]::NewLine,[Text.UTF8Encoding]::new($false))
}
function Save-StateSnapshot {
  $state=[ordered]@{RunId=$script:RunId;CollectorVersion=$script:CollectorVersion;InvocationFingerprint=$script:InvocationFingerprint;UpdatedUtc=[datetime]::UtcNow.ToString('o');CompletedObjects=@($script:CompletedObjects|Sort-Object);Datasets=[ordered]@{}}
  foreach($dataset in $script:PersistedKeys.Keys){$state.Datasets[$dataset]=@($script:PersistedKeys[$dataset])}
  Write-AtomicText $script:StateSnapshotPath ($state|ConvertTo-Json -Depth 8)
}
function Set-ObjectComplete {
  param([string]$Scope,[string]$Name)
  $key="$Scope|$Name"
  if($script:CompletedObjects.Add($key)){Add-CheckpointEvent 'Complete' $Scope $Name;Save-StateSnapshot}
}
function Test-ObjectComplete {
  param([string]$Scope,[string]$Name)
  if($ForceRecollect){return $false}
  $script:CompletedObjects.Contains("$Scope|$Name")
}
function Initialize-DurableState {
  New-Item -ItemType Directory -Force $script:RunStateDir,$script:BatchDir|Out-Null
  $script:InvocationFingerprint=Get-Sha256Text (($script:VCenterNames -join '|')+"|$CollectionProfile|$Days|$IncludePci|$SkipVmFeatures|$VmFeatureSampleLimit|$MaxVMs|$SkipVmStats|ClearTextCredentialFile=$([bool]$ClearTextCredentialFile)")
  if($ResumeRunId){
    if(-not(Test-Path $script:StateSnapshotPath)){throw "Resume state not found: $script:StateSnapshotPath"}
    $state=Get-Content $script:StateSnapshotPath -Raw|ConvertFrom-Json
    if($state.InvocationFingerprint -ne $script:InvocationFingerprint -and -not$ForceRecollect){throw 'Resume invocation does not match the original run. Use identical options or -ForceRecollect.'}
    foreach($key in @($state.CompletedObjects)){[void]$script:CompletedObjects.Add([string]$key)}
    foreach($property in $state.Datasets.PSObject.Properties){
      $set=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
      foreach($key in @($property.Value)){[void]$set.Add([string]$key)}
      $script:PersistedKeys[$property.Name]=$set
    }
    Add-CheckpointEvent 'Resume' 'Run' $script:RunId @{CompletedObjects=$script:CompletedObjects.Count}
  }else{Add-CheckpointEvent 'Start' 'Run' $script:RunId;Save-StateSnapshot}
}
function Add-EvidenceRow {
  param([Parameter(Mandatory)][string]$Dataset,[Parameter(Mandatory)][object]$Row,[string]$SourceVCenter,[string]$SourceHostKey='')
  if(-not $script:DatasetRegistry.Contains($Dataset)){ throw "Unknown dataset: $Dataset" }
  Test-DatasetContract $Dataset $Row
  $evidenceKey=Get-EvidenceKey $Dataset $Row $SourceVCenter $SourceHostKey
  if(-not$script:PersistedKeys.ContainsKey($Dataset)){$script:PersistedKeys[$Dataset]=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)}
  if(-not$script:PendingKeys.ContainsKey($Dataset)){$script:PendingKeys[$Dataset]=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)}
  if($script:PersistedKeys[$Dataset].Contains($evidenceKey) -or $script:PendingKeys[$Dataset].Contains($evidenceKey)){return}
  $meta=[ordered]@{RunId=$script:RunId;EvidenceKey=$evidenceKey;CollectedAtUtc=[datetime]::UtcNow.ToString('o');CollectorVersion=$script:CollectorVersion;SourceVCenter=$SourceVCenter;SourceHostKey=$SourceHostKey}
  foreach($p in $Row.PSObject.Properties){$meta[$p.Name]=$p.Value}
  if(-not $script:Buffers.ContainsKey($Dataset)){$script:Buffers[$Dataset]=[Collections.Generic.List[object]]::new()}
  $script:Buffers[$Dataset].Add([pscustomobject]$meta)
  [void]$script:PendingKeys[$Dataset].Add($evidenceKey)
  if($script:Buffers[$Dataset].Count -ge $FlushRowCount){Flush-EvidenceDataset $Dataset}
}
function Flush-EvidenceDataset {
  param([Parameter(Mandatory)][string]$Dataset)
  if(-not$script:Buffers.ContainsKey($Dataset)){return}
  $rows=@($script:Buffers[$Dataset]);if($rows.Count -eq 0){return}
  $path=Join-Path $script:OutDir $script:DatasetRegistry[$Dataset].File
  $batchPath=Join-Path $script:BatchDir ("{0}-{1}-{2}.csv" -f $Dataset,[datetime]::UtcNow.ToString('yyyyMMddHHmmssfff'),[guid]::NewGuid().ToString('N'))
  $rows|Export-Csv -LiteralPath $batchPath -NoTypeInformation -Encoding utf8
  # Batch is durable before destination append. EvidenceKey makes replay idempotent.
  if(Test-Path $path){
    $reader=[IO.File]::OpenText($batchPath);try{[void]$reader.ReadLine();$writer=[IO.StreamWriter]::new($path,$true,[Text.UTF8Encoding]::new($false));try{while(($line=$reader.ReadLine()) -ne $null){$writer.WriteLine($line)}}finally{$writer.Flush();$writer.Dispose()}}finally{$reader.Dispose()}
  }else{Copy-Item $batchPath $path}
  foreach($row in $rows){[void]$script:PersistedKeys[$Dataset].Add([string]$row.EvidenceKey)}
  $script:Buffers[$Dataset].Clear();$script:PendingKeys[$Dataset].Clear()
  $script:Metrics.RowsCommitted+=$rows.Count;$script:Metrics.BatchesCommitted++;$script:LastCheckpointUtc=[datetime]::UtcNow.ToString('o')
  Add-CheckpointEvent 'BatchCommitted' 'Dataset' $Dataset @{Rows=$rows.Count;Batch=$batchPath;Destination=$path}
  Write-CollectorEvent CHECKPOINT "Committed $($rows.Count) row(s)" BatchCommitted $Dataset '' '' $Dataset 'Flush-EvidenceDataset' 0 $rows.Count
  Save-StateSnapshot
}
function Flush-AllEvidence { foreach($d in @($script:Buffers.Keys)){Flush-EvidenceDataset $d} }
function Complete-RequestedDatasets {
  foreach($dataset in Get-ActiveDatasetNames){
    $path=Join-Path $script:OutDir $script:DatasetRegistry[$dataset].File
    if(Test-Path -LiteralPath $path){
      $rows=[math]::Max(0,@(Get-Content -LiteralPath $path).Count-1)
      $script:DatasetStatus[$dataset]=[ordered]@{Status='Produced';Rows=$rows;File=$path}
      Set-DatasetStatus $dataset Produced '' '' $rows 0
    }else{
      if(-not $script:DatasetTelemetry.Contains($dataset)){Set-DatasetStatus $dataset NotImplemented 'Requested dataset produced no rows and no collector query registered.' '' 0 0}
      $latest=@($script:DatasetTelemetry[$dataset])[-1]
      $script:DatasetStatus[$dataset]=[ordered]@{Status=$latest.Status;Rows=0;File='';Error=$latest.Error}
      Write-CollectorEvent WARN "No inventory CSV created for $dataset; status recorded in dataset-status.json as $($latest.Status)" DatasetGap $dataset '' '' $dataset 'Complete-RequestedDatasets'
    }
  }
  $contract=[ordered]@{SchemaVersion='4.1';RunId=$script:RunId;CollectorVersion=$script:CollectorVersion;GeneratedUtc=[datetime]::UtcNow.ToString('o');Datasets=$script:DatasetTelemetry}
  Write-AtomicText $script:DatasetStatusPath ($contract|ConvertTo-Json -Depth 10)
}
function Write-RunManifest {
  param([ValidateSet('Completed','CompletedWithGaps','CompletedWithErrors','Failed')][string]$Status,[string]$Error='')
  $manifest=[ordered]@{RunId=$script:RunId;CollectorVersion=$script:CollectorVersion;Profile=$CollectionProfile;StartedUtc=$script:RunStartedUtc.ToString('o');UpdatedUtc=[datetime]::UtcNow.ToString('o');Status=$Status;VCentersRequested=@($script:VCenterNames);RequestedDatasets=@(Get-ActiveDatasetNames);DatasetStatus=$script:DatasetStatus;Error=$Error}
  $manifest|ConvertTo-Json -Depth 8|Set-Content -LiteralPath $script:ManifestPath -Encoding utf8
  if($Status -in @('Completed','CompletedWithGaps')){$manifest|ConvertTo-Json -Depth 8|Set-Content -LiteralPath (Join-Path $script:OutDir 'latest-completed-run.json') -Encoding utf8}
}
function Save-Checkpoint {
  param([string]$Phase,[string]$VCenterName,[string]$ObjectName='')
  if($DryRun){return}
  [ordered]@{RunId=$script:RunId;CollectorVersion=$script:CollectorVersion;UpdatedUtc=[datetime]::UtcNow.ToString('o');Phase=$Phase;VCenter=$VCenterName;Object=$ObjectName}|ConvertTo-Json|Set-Content -LiteralPath $script:CheckpointPath -Encoding utf8
}
function Get-Percentile {param([double[]]$Values,[double]$P) if(-not $Values -or $Values.Count-eq 0){return $null};$s=@($Values|Sort-Object);$i=[math]::Min($s.Count-1,[math]::Ceiling($P*$s.Count)-1);$s[[math]::Max($i,0)]}
function Get-SafeProperty {param($InputObject,[string]$Name,$Default=$null) if($null-eq $InputObject){return $Default};$p=$InputObject.PSObject.Properties[$Name];if($p){$p.Value}else{$Default}}
function ConvertTo-NullableBoolean {
  param($Value)
  if($null -eq $Value){return $null}
  if($Value -is [bool]){return $Value}
  switch -Regex (([string]$Value).Trim()) {
    '^(?i:true|1|yes|enabled)$' {return $true}
    '^(?i:false|0|no|disabled)$' {return $false}
    default {return $null}
  }
}
function ConvertTo-PciHex4 {
  param($Value)
  if($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)){return ''}
  $raw=([string]$Value).Trim().ToLowerInvariant()
  if($raw -match '^0x([0-9a-f]{1,4})$'){return $Matches[1].PadLeft(4,'0')}
  if($raw -match '^[0-9a-f]{1,4}$' -and $raw -match '[a-f]'){return $raw.PadLeft(4,'0')}
  $number=0L
  if([long]::TryParse($raw,[ref]$number)){return ('{0:x4}' -f $number)}
  return $raw
}
function Get-SafePath {
  param($InputObject,[Parameter(Mandatory)][string[]]$Path,$Default=$null)
  $value=$InputObject
  foreach($part in $Path){
    if($null-eq$value){return $Default}
    $property=$value.PSObject.Properties[$part]
    if(-not$property){return $Default}
    $value=$property.Value
  }
  if($null-eq$value){$Default}else{$value}
}
function Assert-Prerequisites {
  if($DryRun){return}
  if(-not(Get-Module -ListAvailable VMware.VimAutomation.Core)){throw 'VMware.PowerCLI 13.x or later is required.'}
  Import-Module VMware.VimAutomation.Core -ErrorAction Stop
  Set-PowerCLIConfiguration -Scope Session -ParticipateInCEIP:$false -Confirm:$false|Out-Null
  Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Confirm:$false -Scope Session|Out-Null
}
function Get-VCenterNames {
  param([Parameter(Mandatory)][string]$Path)
  if(-not (Test-Path -LiteralPath $Path -PathType Leaf)){throw "vCenter list file not found: $Path"}
  $seen=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
  $items=[Collections.Generic.List[string]]::new()
  foreach($line in Get-Content -LiteralPath $Path -ErrorAction Stop){
    $value=([string]$line).Trim()
    if(-not $value -or $value.StartsWith('#')){continue}
    # Permit an end-of-line comment after whitespace; preserve valid hyphens and dots.
    $value=($value -split '\s+#',2)[0].Trim().TrimEnd('.')
    if($value -and $seen.Add($value)){$items.Add($value)}
  }
  if($items.Count -eq 0){throw "No vCenter names found in $Path"}
  @($items)
}
function Import-CredentialMap {
  param([string]$Path)
  $map=@{}
  if(-not $Path){return $map}
  if(-not(Test-Path -LiteralPath $Path -PathType Leaf)){throw "Credential map not found: $Path"}
  $source=Import-Clixml -LiteralPath $Path
  if($source -is [Collections.IDictionary]){
    foreach($key in $source.Keys){if($source[$key] -isnot [pscredential]){throw "Credential map entry '$key' is not a PSCredential"};$map[[string]$key]=$source[$key]}
  } else {
    foreach($entry in @($source)){
      $name=[string](Get-SafeProperty $entry 'VCenter')
      $cred=Get-SafeProperty $entry 'Credential'
      if(-not$name -or $cred -isnot [pscredential]){throw 'Credential map entries require VCenter and Credential (PSCredential)'}
      $map[$name]=$cred
    }
  }
  $map
}
function Read-VCenterCredential {
  param([Parameter(Mandatory)][string]$Message)
  # Passing -UserName changes Get-Credential into a password-only prompt. Leave
  # it unset by default so both user name and password are requested.
  if([string]::IsNullOrWhiteSpace($DefaultUserName)){
    Get-Credential -Message $Message
  } else {
    Get-Credential -UserName $DefaultUserName -Message $Message
  }
}
function Import-ClearTextCredentialFile {
  param([string]$Path)
  $map=@{}
  if([string]::IsNullOrWhiteSpace($Path)){return $map}
  if(-not$AllowClearTextCredentials){
    throw 'Clear-text credential file requested. Re-run with -AllowClearTextCredentials to acknowledge the risk.'
  }
  if(-not(Test-Path -LiteralPath $Path -PathType Leaf)){throw "Clear-text credential file not found: $Path"}
  $item=Get-Item -LiteralPath $Path
  if($item.Length -eq 0){throw "Clear-text credential file is empty: $Path"}
  $rows=@(Import-Csv -LiteralPath $Path)
  if(-not$rows.Count){throw "Clear-text credential file contains no rows: $Path"}
  $required=@('vCenterFqdn','Username','Password')
  foreach($column in $required){
    if(-not($rows[0].PSObject.Properties.Name -contains $column)){throw "Clear-text credential file is missing required column: $column"}
  }
  foreach($row in $rows){
    $name=([string]$row.vCenterFqdn).Trim().TrimEnd('.')
    $user=([string]$row.Username).Trim()
    $password=[string]$row.Password
    if(-not$name -or -not$user -or [string]::IsNullOrEmpty($password)){throw 'Every credential row requires vCenterFqdn, Username, and Password'}
    if($map.ContainsKey($name)){throw "Duplicate vCenterFqdn in credential file: $name"}
    $secure=ConvertTo-SecureString $password -AsPlainText -Force
    $map[$name]=[pscredential]::new($user,$secure)
    # Remove the plaintext value from the imported object as soon as the PSCredential exists.
    $row.Password=$null;$password=$null
  }
  Write-CollectorEvent WARN "Loaded $($map.Count) credential mapping(s) from an explicitly approved clear-text file. File contents were not logged." CredentialFileLoaded $Path
  $map
}
function Initialize-Credentials {
  param([string[]]$Names)
  $imported=Import-CredentialMap $CredentialMap
  foreach($key in $imported.Keys){$script:Credentials[$key]=$imported[$key]}
  $clearText=Import-ClearTextCredentialFile $ClearTextCredentialFile
  foreach($key in $clearText.Keys){
    if($script:Credentials.ContainsKey($key)){throw "Credential for '$key' was supplied by both CredentialMap and ClearTextCredentialFile"}
    $script:Credentials[$key]=$clearText[$key]
  }
  $unknown=@($script:Credentials.Keys|Where-Object{$_ -notin $Names})
  if($unknown.Count){throw "Credential file/map contains vCenter(s) not present in VCenterList: $($unknown -join ', ')"}
  $missing=@($Names|Where-Object{-not$script:Credentials.ContainsKey($_)})
  if($missing.Count -eq 0){return}
  if($SharedCredential){
    $cred=Read-VCenterCredential -Message "Shared read-only credential for $($missing.Count) vCenter(s)"
    if(-not$cred){throw 'Shared credential acquisition was canceled'}
    foreach($name in $missing){$script:Credentials[$name]=$cred}
  } else {
    # Acquire every credential before any connection or collection starts.
    foreach($name in $missing){
      $cred=Read-VCenterCredential -Message "Read-only credential for $name"
      if(-not$cred){throw "Credential acquisition was canceled for $name"}
      $script:Credentials[$name]=$cred
    }
  }
}
function Invoke-ConnectionPreflight {
  param([string[]]$Names)
  foreach($name in $Names){
    $started=[datetime]::UtcNow
    try{
      Write-CollectorLog INFO 'Connection preflight started' $name
      $server=Invoke-WithRetry -Operation 'Connect-VIServer' -Target $name -VCenterName $name -ScriptBlock { Connect-VIServer $name -Credential $script:Credentials[$name] -NotDefault -ErrorAction Stop }
      # Require usable About data so a TCP-only or partially authorized connection is not accepted.
      $about=$server.ExtensionData.Content.About
      if(-not$about.Version){throw 'Connected, but vCenter About.Version was unavailable'}
      $script:PreflightServers[$name]=$server
      $script:PreflightStatus.Add([pscustomobject]@{VCenter=$name;Status='Ready';Version=$about.Version;Build=$about.Build;ElapsedSeconds=[math]::Round(([datetime]::UtcNow-$started).TotalSeconds,2);Error=''})
      Write-CollectorLog INFO "Connection preflight passed version=$($about.Version) build=$($about.Build)" $name
    } catch {
      $script:PreflightStatus.Add([pscustomobject]@{VCenter=$name;Status='Failed';Version='';Build='';ElapsedSeconds=[math]::Round(([datetime]::UtcNow-$started).TotalSeconds,2);Error=$_.Exception.Message})
      Write-CollectorLog ERROR "Connection preflight failed: $($_.Exception.Message)" $name
    }
  }
  if(-not$DryRun){
    $statusPath=Join-Path $script:OutDir 'connection-preflight.csv'
    $script:PreflightStatus|Export-Csv -LiteralPath $statusPath -NoTypeInformation -Encoding utf8
  }
  $failed=@($script:PreflightStatus|Where-Object Status -eq 'Failed')
  if($failed.Count -and $ConnectionFailureAction -eq 'Stop'){
    foreach($server in $script:PreflightServers.Values){Disconnect-VIServer $server -Confirm:$false -ErrorAction SilentlyContinue}
    $script:PreflightServers.Clear()
    throw "Connection preflight failed for $($failed.Count) vCenter(s); no collection started"
  }
  if($script:PreflightServers.Count -eq 0){throw 'No vCenter passed connection preflight; no collection started'}
}


function Invoke-CoreInventory {
  param($Server,[string]$VcName)
  $script:Metrics.Operations++;Write-CollectorEvent INFO 'Core inventory started' OperationStarted $VcName $VcName '' '' 'Invoke-CoreInventory'
  $about=$Server.ExtensionData.Content.About
  Add-EvidenceRow pcli_vcenter ([pscustomobject]@{vCenter=$VcName;Version=$about.Version;Build=$about.Build;FullName=$about.FullName;InstanceUuid=$about.InstanceUuid;NsxRegistered=$null;VCHA='NotQueriedInWorkItem1';LinkedVCenters=[math]::Max(0,($script:PreflightServers.Count-1))}) $VcName
  foreach($c in Get-Cluster -Server $Server){
    $cx=$c.ExtensionData;$vsc=$cx.ConfigurationEx.VsanConfigInfo
    Add-EvidenceRow pcli_clusters ([pscustomobject]@{vCenter=$VcName;Cluster=$c.Name;Hosts=@(Get-VMHost -Location $c).Count;DrsEnabled=$c.DrsEnabled;DrsLevel=$c.DrsAutomationLevel;HAEnabled=$c.HAEnabled;EVC=$c.EVCMode;LifecycleMode='NotQueriedInWorkItem1';VsanEnabled=[bool]$vsc.Enabled;VsanESA=(Get-SafeProperty $vsc 'VsanEsaEnabled');SupervisorStatus='NotQueriedInWorkItem1'}) $VcName
  }
  foreach($d in Get-Datastore -Server $Server){
    # PowerCLI versions expose free capacity as FreeSpaceGB, ExtensionData.Summary.FreeSpace,
    # or occasionally neither. Never let one missing convenience property abort the vCenter.
    $capacityGB=Get-SafeProperty $d 'CapacityGB'
    if($null -eq $capacityGB -and $d.ExtensionData.Summary.Capacity){$capacityGB=$d.ExtensionData.Summary.Capacity/1GB}
    $freeGB=Get-SafeProperty $d 'FreeSpaceGB'
    if($null -eq $freeGB){$freeGB=Get-SafeProperty $d 'FreeGB'}
    if($null -eq $freeGB -and $null -ne $d.ExtensionData.Summary.FreeSpace){$freeGB=$d.ExtensionData.Summary.FreeSpace/1GB}
    $fsVersion=Get-SafeProperty $d 'FileSystemVersion'
    Add-EvidenceRow pcli_datastores ([pscustomobject]@{
      vCenter=$VcName;Name=$d.Name;Type=$d.Type
      CapacityGB=$(if($null-ne$capacityGB){[math]::Round([double]$capacityGB,1)})
      FreeGB=$(if($null-ne$freeGB){[math]::Round([double]$freeGB,1)})
      VmfsVersion=$fsVersion;HostCount=@($d.ExtensionData.Host).Count
      FreeSpaceEvidence=$(if($null-ne$freeGB){'Direct'}else{'NotExposed'})
      CapacitySource=$(if($null-ne(Get-SafeProperty $d 'CapacityGB')){'PowerCLI.CapacityGB'}elseif($null-ne(Get-SafePath $d.ExtensionData @('Summary','Capacity') $null)){'ExtensionData.Summary.CapacityBytes'}else{'NotExposed'})
      FreeSpaceSource=$(if($null-ne(Get-SafeProperty $d 'FreeSpaceGB')){'PowerCLI.FreeSpaceGB'}elseif($null-ne(Get-SafeProperty $d 'FreeGB')){'PowerCLI.FreeGB'}elseif($null-ne(Get-SafePath $d.ExtensionData @('Summary','FreeSpace') $null)){'ExtensionData.Summary.FreeSpaceBytes'}else{'NotExposed'})
      CapacityValidation=$(if($null-eq$capacityGB){'NotExposed'}elseif($capacityGB-lt 0){'InvalidNegative'}elseif($freeGB-gt$capacityGB){'FreeExceedsCapacity'}elseif($capacityGB-gt 262144){'VerifyLargeCapacity'}else{'Plausible'})
    }) $VcName
  }
  foreach($d in Get-VDSwitch -Server $Server){Add-EvidenceRow pcli_vds ([pscustomobject]@{vCenter=$VcName;Name=$d.Name;Version=$d.Version;Mtu=$d.Mtu;NumUplinkPorts=$d.NumUplinkPorts;HostCount=@($d.ExtensionData.Summary.HostMember).Count;LacpEnabled=($null-ne$d.ExtensionData.Config.LacpApiVersion);Vendor=$d.Vendor}) $VcName}
  $allHosts=@(Get-VMHost -Server $Server);$hostIndex=0
  foreach($h in $allHosts){$hostIndex++;if(Test-ObjectComplete 'HostCore' "$VcName|$($h.Name)"){Write-CollectorEvent PROGRESS 'Resume skip: host core complete' HostSkipped $h.Name $VcName $h.Name '' 'CoreHostLoop';continue};Write-CollectorEvent PROGRESS "Host $hostIndex of $($allHosts.Count)" HostProgress $h.Name $VcName $h.Name '' 'CoreHostLoop'
    $hx=$h.ExtensionData;$cpuFam=$null;$cpuMod=$null;$cpuStep=$null;$x=$null
    if($h.ConnectionState -eq 'Connected'){try{$x=Get-EsxCli -VMHost $h -V2;$cl=$x.hardware.cpu.list.Invoke()|Select-Object -First 1;$cpuFam=$cl.Family;$cpuMod=$cl.Model;$cpuStep=$cl.Stepping}catch{Write-CollectorLog WARN "CPU ESXCLI failed: $($_.Exception.Message)" $h.Name}}
    Add-EvidenceRow pcli_hosts ([pscustomobject]@{vCenter=$VcName;Host=$h.Name;InventoryNameIsFqdn=($h.Name-match'\.');Cluster=$h.Parent.Name;Vendor=$hx.Hardware.SystemInfo.Vendor;Model=$hx.Hardware.SystemInfo.Model;Bios=$hx.Hardware.BiosInfo.BiosVersion;CpuModel=$(if($hx.Hardware.CpuPkg.Count){$hx.Hardware.CpuPkg[0].Description});CpuFamily=$cpuFam;CpuModelNum=$cpuMod;CpuStepping=$cpuStep;Tpm=$null;EsxVersion=$h.Version;EsxBuild=$h.Build;Connection=$h.ConnectionState;LockdownMode=(Get-SafePath $hx @('Config','LockdownMode') 'NotExposed')}) $VcName $h.Name
    # Get-VMHostNetworkAdapter is valid only for Connected or Maintenance hosts.
    # Record inventory for every host, but isolate unavailable VMkernel evidence.
    if($h.ConnectionState -in @('Connected','Maintenance')){
      try{
        foreach($n in Get-VMHostNetworkAdapter -VMHost $h -VMKernel -ErrorAction Stop){
          Add-EvidenceRow pcli_vmk ([pscustomobject]@{
            vCenter=$VcName;Host=$h.Name;Cluster=$h.Parent.Name;Device=$n.Name
            PortGroup=$n.PortGroupName;OnVds=[bool](Get-SafePath $n.ExtensionData @('Spec','DistributedVirtualPort') $null)
            IP=$n.IP;Dhcp=$n.DhcpEnabled;Mtu=$n.Mtu;Mgmt=$n.ManagementTrafficEnabled
            vMotion=$n.VMotionEnabled;FT=$n.FaultToleranceLoggingEnabled
            vSAN=$n.VsanTrafficEnabled;Provisioning=$n.ProvisioningEnabled
            CollectionStatus='Direct';CollectionError=''
          }) $VcName $h.Name
        }
      } catch {
        Write-CollectorLog WARN "VMkernel adapter query failed; host collection continues: $($_.Exception.Message)" $h.Name
        Add-EvidenceRow pcli_vmk ([pscustomobject]@{
          vCenter=$VcName;Host=$h.Name;Cluster=$h.Parent.Name;Device='';PortGroup=''
          OnVds=$null;IP='';Dhcp=$null;Mtu=$null;Mgmt=$null;vMotion=$null;FT=$null
          vSAN=$null;Provisioning=$null;CollectionStatus='QueryFailed';CollectionError=$_.Exception.Message
        }) $VcName $h.Name
      }
    } else {
      Write-CollectorEvent WARN "Skipping VMkernel adapter query because host state is $($h.ConnectionState)" UnavailableObject $h.Name $VcName $h.Name 'pcli_vmk' 'Get-VMHostNetworkAdapter'
      Add-EvidenceRow pcli_vmk ([pscustomobject]@{
        vCenter=$VcName;Host=$h.Name;Cluster=$h.Parent.Name;Device='';PortGroup=''
        OnVds=$null;IP='';Dhcp=$null;Mtu=$null;Mgmt=$null;vMotion=$null;FT=$null
        vSAN=$null;Provisioning=$null;CollectionStatus='NotAvailable'
        CollectionError="Host state $($h.ConnectionState); requires Connected or Maintenance"
      }) $VcName $h.Name
    }
    if($IncludePci -and $x){try{foreach($p in $x.hardware.pci.list.Invoke()){if($p.ModuleName -and ([string]$p.DeviceClassName-match'Network|Ethernet|Fibre|Mass storage|RAID|SCSI|NVM|SATA|SAS|Serial Attached')){Add-EvidenceRow pcli_pci ([pscustomobject]@{vCenter=$VcName;Host=$h.Name;Cluster=$h.Parent.Name;Device=$p.DeviceName;Vendor=$p.VendorName;VID=(ConvertTo-PciHex4 $p.VendorID);DID=(ConvertTo-PciHex4 $p.DeviceID);SVID=(ConvertTo-PciHex4 $p.SubVendorID);SSID=(ConvertTo-PciHex4 $p.SubDeviceID);Driver=$p.ModuleName;Class=$p.DeviceClassName}) $VcName $h.Name}}}catch{Write-CollectorLog WARN "PCI query failed; host collection continues: $($_.Exception.Message)" $h.Name}}
    Flush-AllEvidence;Set-ObjectComplete 'HostCore' "$VcName|$($h.Name)";if(($hostIndex%$HostBatchSize)-eq 0){Write-CollectorEvent PROGRESS "Host core checkpoint $hostIndex of $($allHosts.Count)" Heartbeat $h.Name $VcName $h.Name '' 'CoreHostLoop'}
  }
}
function Invoke-VmEvidence {
  param($Server,[string]$VcName)
  $started=[datetime]::UtcNow
  $props='Name','Runtime.PowerState','Config.Template','Config.Uuid','Config.KeyId','Config.Firmware','Config.BootOptions','Config.Flags','Config.Hardware.Device','Runtime.Host'
  $views=@(Get-View -Server $Server -ViewType VirtualMachine -Property $props|Where-Object{-not$_.Config.Template}|Sort-Object Name,{$_.Config.Uuid})
  $vmIndex=0;foreach($v in $views){$vmIndex++;if(Test-ObjectComplete 'VMFeature' "$VcName|$($v.Config.Uuid)"){continue};if(($vmIndex%100)-eq 0){Write-CollectorEvent PROGRESS "VM evidence $vmIndex of $($views.Count)" Heartbeat $v.Name $VcName '' '' 'VmEvidence'}
    $dev=@($v.Config.Hardware.Device);$encDisks=@($dev|Where-Object{$_-is[VMware.Vim.VirtualDisk]-and$_.Backing.KeyId}).Count;$kp=Get-SafePath $v @('Config','KeyId','KeyId') ''
    Add-EvidenceRow pcli_vm_encryption ([pscustomobject]@{vCenter=$VcName;VM=$v.Name;Uuid=$v.Config.Uuid;Encryption=$(if($v.Config.KeyId-or$encDisks-gt 0){'Yes'}else{'No'});VmConfigEncrypted=[bool]$v.Config.KeyId;EncryptedDisks=$encDisks;KeyProviderId=$kp;PowerState=$v.Runtime.PowerState;CollectionStatus='Direct';CollectionError=''}) $VcName
    $vbs=Get-SafePath $v @('Config','Flags','VbsEnabled') $null;$iommu=@($dev|Where-Object{$_.GetType().Name-match'Iommu'}).Count-gt 0
    Add-EvidenceRow pcli_vm_features ([pscustomobject]@{vCenter=$VcName;VM=$v.Name;Uuid=$v.Config.Uuid;Firmware=$v.Config.Firmware;EfiSecureBoot=(Get-SafePath $v @('Config','BootOptions','EfiSecureBootEnabled') $null);VBS=$vbs;IOMMU=$iommu;vTPM=@($dev|Where-Object{$_-is[VMware.Vim.VirtualTPM]}).Count-gt 0;PciPassthrough=@($dev|Where-Object{$_-is[VMware.Vim.VirtualPCIPassthrough]}).Count;PowerState=$v.Runtime.PowerState;CollectionStatus='Direct';CollectionError=''}) $VcName
    foreach($disk in @($dev|Where-Object{$_-is[VMware.Vim.VirtualDisk]})){
      $sharing=[string](Get-SafePath $disk @('Backing','Sharing') '');$bus='';$controller=@($dev|Where-Object{$_.Key-eq$disk.ControllerKey}|Select-Object -First 1);if($controller){$bus=[string](Get-SafeProperty $controller[0] 'SharedBus' '')}
      if($sharing-match'multiwriter|sharingMultiWriter' -or $bus-match'physicalSharing|virtualSharing'){
        Add-EvidenceRow pcli_vm_shared_disks ([pscustomobject]@{vCenter=$VcName;VM=$v.Name;Uuid=$v.Config.Uuid;DiskLabel=$disk.DeviceInfo.Label;ControllerBusSharing=$bus;SharingMode=$sharing;BackingFile=(Get-SafePath $disk @('Backing','FileName') '');CollectionStatus='Direct';CollectionError=''}) $VcName
      }
    }
    Flush-AllEvidence;Set-ObjectComplete 'VMFeature' "$VcName|$($v.Config.Uuid)"
  }
  Add-EvidenceRow pcli_vm_coverage ([pscustomobject]@{vCenter=$VcName;TotalNonTemplateVMs=$views.Count;EncryptionRows=$views.Count;DetailedFeatureRows=$views.Count;DetailedFeatureLimit=0;Selection='All non-template VMs; WI26 full coverage'}) $VcName
  Set-DatasetStatus pcli_vm_features Produced '' $VcName $views.Count (([datetime]::UtcNow-$started).TotalSeconds)
  Set-DatasetStatus pcli_vm_encryption Produced '' $VcName $views.Count (([datetime]::UtcNow-$started).TotalSeconds)
}
function Invoke-HostDeepDive {
  param($Server,[string]$VcName)
  $script:Metrics.Operations++;Write-CollectorEvent INFO 'Host deep dive started' OperationStarted $VcName $VcName '' '' 'Invoke-HostDeepDive'
  foreach($h in Get-VMHost -Server $Server|Where-Object ConnectionState -eq 'Connected'){
    $hn=$h.Name;if(Test-ObjectComplete 'HostDeepDive' "$VcName|$hn"){continue};$cl=$h.Parent.Name;$fwd=$false;$rev=$false;$ip=''
    try{$ip=([Net.Dns]::GetHostAddresses($hn)|Select-Object -First 1).IPAddressToString;$fwd=$true;$rev=([Net.Dns]::GetHostEntry($ip).HostName-ieq$hn)}catch{}
    Add-EvidenceRow pcli_dns ([pscustomobject]@{vCenter=$VcName;Cluster=$cl;Host=$hn;InventoryNameIsFqdn=($hn-match'\.');ForwardOk=$fwd;ReverseOk=$rev;ResolvedIP=$ip}) $VcName $hn
    $svc=Get-VMHostService -VMHost $h|Where-Object Key -eq 'ntpd';Add-EvidenceRow pcli_ntp ([pscustomobject]@{vCenter=$VcName;Cluster=$cl;Host=$hn;NtpServers=((Get-VMHostNtpServer -VMHost $h)-join';');NtpRunning=$svc.Running;NtpPolicy=$svc.Policy}) $VcName $hn
    try{$x=Get-EsxCli -VMHost $h -V2
      foreach($dv in $x.storage.core.device.list.Invoke()){
        $isBoot=ConvertTo-NullableBoolean $dv.IsBootDevice
        $isBootUsb=ConvertTo-NullableBoolean $dv.IsBootUSBDevice
        if($isBoot -eq $true -or $isBootUsb -eq $true){
          Add-EvidenceRow pcli_boot ([pscustomobject]@{
            vCenter=$VcName;Cluster=$cl;Host=$hn;Device=$dv.Device;Vendor=$dv.Vendor;Model=$dv.Model;SizeMB=$dv.Size
            IsBootDevice=$isBoot;IsBootUSB=$isBootUsb;IsSSD=(ConvertTo-NullableBoolean $dv.IsSSD)
            IsLocal=(ConvertTo-NullableBoolean $dv.IsLocal);IsRemovable=(ConvertTo-NullableBoolean $dv.IsRemovable)
            CollectionStatus='Direct';CollectionError=''
          }) $VcName $hn
        }
      }
      $enc=$x.system.settings.encryption.get.Invoke();Add-EvidenceRow pcli_security ([pscustomobject]@{vCenter=$VcName;Cluster=$cl;Host=$hn;EncryptionMode=(Get-SafeProperty $enc 'Mode' 'NotExposed');RequireSecureBoot=(Get-SafeProperty $enc 'RequireSecureBoot' 'NotExposed');SecureBoot=(Get-SafePath $h.ExtensionData @('Capability','TpmSupported') 'NotExposed');TpmPresent=(Get-SafePath $h.ExtensionData @('Hardware','TpmInfo') $null)-ne$null;TpmVersion=(Get-SafePath $h.ExtensionData @('Hardware','TpmInfo','Version') 'NotExposed');TpmMode=(Get-SafePath $h.ExtensionData @('Hardware','TpmInfo','Mode') 'NotExposed');Attestation=(Get-SafePath $h.ExtensionData @('Runtime','TpmPcrValues') 'NotExposed');LockdownMode=(Get-SafePath $h.ExtensionData @('Config','LockdownMode') 'NotExposed');SshRunning=[bool]((Get-VMHostService -VMHost $h|Where-Object Key -eq 'TSM-SSH').Running);CertificateStatus=(Get-SafePath $h.ExtensionData @('Runtime','HealthSystemRuntime','SystemHealthInfo','NumericSensorInfo') 'NotExposed');CollectionStatus='Direct';CollectionError=''}) $VcName $hn
      foreach($v in $x.software.vib.list.Invoke()){if($v.Vendor-notmatch'^VMware|VMW$'){Add-EvidenceRow pcli_vibs ([pscustomobject]@{vCenter=$VcName;Cluster=$cl;Host=$hn;Name=$v.Name;Version=$v.Version;Vendor=$v.Vendor;AcceptanceLevel=$v.AcceptanceLevel;InstallDate=$v.InstallDate}) $VcName $hn}}
      foreach($n in $x.network.nic.list.Invoke()){$g=$null;try{$g=$x.network.nic.get.Invoke(@{nicname=$n.Name})}catch{};$di=Get-SafeProperty $g 'DriverInfo';Add-EvidenceRow pcli_nicfw ([pscustomobject]@{vCenter=$VcName;Cluster=$cl;Host=$hn;Nic=$n.Name;Description=$n.Description;SpeedMbps=(Get-SafeProperty $n 'LinkSpeed');Driver=(Get-SafeProperty $di 'Driver');DriverVersion=(Get-SafeProperty $di 'Version');Firmware=(Get-SafeProperty $di 'FirmwareVersion');PCI=(Get-SafeProperty $n 'PCIDevice')}) $VcName $hn}
    }catch{Write-CollectorLog WARN "Deep-dive ESXCLI failed: $($_.Exception.Message)" $hn}
    Flush-AllEvidence;Set-ObjectComplete 'HostDeepDive' "$VcName|$hn"
  }
}
function Get-AvailableMemoryCounters {
  param($Entity,[string]$Target)
  try {
    $all=@(Get-StatType -Entity $Entity -ErrorAction Stop | ForEach-Object {[string]$_} | Sort-Object -Unique)
    $memory=@($all | Where-Object {$_ -match '^(?i:mem\.)'})
    [pscustomobject]@{Status='Direct';Error='';All=$memory}
  } catch {
    $message=$_.Exception.Message
    Write-CollectorLog WARN "Memory counter discovery failed: $message" $Target
    [pscustomobject]@{Status='QueryFailed';Error=$message;All=@()}
  }
}
function Select-ActiveMemoryCounter {
  param([string[]]$Available)
  # Preserve semantic integrity: only select counters that represent active memory.
  # Do not silently substitute consumed, granted, usage, or configured memory.
  $preference=@(
    'mem.active.average',
    'mem.active',
    'mem.active.latest',
    'mem.active.maximum',
    'mem.active.minimum'
  )
  foreach($wanted in $preference){
    $hit=@($Available | Where-Object {$_ -ieq $wanted} | Select-Object -First 1)
    if($hit.Count){return [string]$hit[0]}
  }
  # Permit namespace variants only when the terminal metric is explicitly active.
  $pattern=@($Available | Where-Object {$_ -match '^(?i:mem\.active(?:\.|$))'} | Sort-Object | Select-Object -First 1)
  if($pattern.Count){return [string]$pattern[0]}
  return $null
}
function Get-MemoryStatSamples {
  param($Entity,[datetime]$Start,[string]$Target)
  $requested='mem.active.average'
  $discovery=Get-AvailableMemoryCounters $Entity $Target
  $availableText=(@($discovery.All) -join ';')
  if($discovery.Status -ne 'Direct'){
    return [pscustomobject]@{
      Status='QueryFailed';Error="Counter discovery failed: $($discovery.Error)"
      RequestedMetric=$requested;SelectedMetric='';AvailableMemoryMetrics=$availableText
      Interval='';Stats=@();First=$null;Last=$null
    }
  }
  $selected=Select-ActiveMemoryCounter $discovery.All
  if([string]::IsNullOrWhiteSpace($selected)){
    $message="No active-memory counter is exposed. Available memory counters: $availableText"
    Write-CollectorEvent WARN $message EvidenceGap $Target '' '' 'memory' 'Get-StatType'
    return [pscustomobject]@{
      Status='NotExposed';Error=$message;RequestedMetric=$requested;SelectedMetric=''
      AvailableMemoryMetrics=$availableText;Interval='';Stats=@();First=$null;Last=$null
    }
  }
  Write-CollectorLog INFO "Memory counter selected: requested=$requested selected=$selected" $Target
  $attempts=@(
    @{IntervalMins=30;Label='30-minute'},
    @{IntervalMins=120;Label='2-hour'},
    @{IntervalMins=1440;Label='daily'},
    @{IntervalMins=$null;Label='automatic'}
  )
  $errors=[Collections.Generic.List[string]]::new()
  foreach($attempt in $attempts){
    try {
      $params=@{Entity=$Entity;Stat=$selected;Start=$Start;ErrorAction='Stop'}
      if($null-ne$attempt.IntervalMins){$params.IntervalMins=$attempt.IntervalMins}
      $stats=@(Get-Stat @params | Sort-Object Timestamp)
      if($stats.Count){
        return [pscustomobject]@{
          Status='Direct';Error='';RequestedMetric=$requested;SelectedMetric=$selected
          AvailableMemoryMetrics=$availableText;Interval=$attempt.Label;Stats=$stats
          First=$stats[0].Timestamp;Last=$stats[-1].Timestamp
        }
      }
      $errors.Add("$($attempt.Label): no samples")
    } catch {$errors.Add("$($attempt.Label): $($_.Exception.Message)")}
  }
  $message="Selected counter '$selected' returned no usable samples: $($errors -join ' | ')"
  Write-CollectorLog WARN $message $Target
  [pscustomobject]@{
    Status='QueryFailed';Error=$message;RequestedMetric=$requested;SelectedMetric=$selected
    AvailableMemoryMetrics=$availableText;Interval='';Stats=@();First=$null;Last=$null
  }
}
function Invoke-MemoryEvidence {
  param($Server,[string]$VcName)
  $script:Metrics.Operations++;Write-CollectorEvent INFO 'Memory evidence started' OperationStarted $VcName $VcName '' 'memory' 'Invoke-MemoryEvidence'
  $start=(Get-Date).AddDays(-$Days);$hosts=Get-VMHost -Server $Server|Where-Object ConnectionState -eq 'Connected'
  foreach($h in $hosts){
    if(Test-ObjectComplete 'HostMemory' "$VcName|$($h.Name)"){continue}
    $result=Get-MemoryStatSamples $h $start $h.Name;$act=@($result.Stats|ForEach-Object{$_.Value/1MB})
    Add-EvidenceRow pcli_mem_host ([pscustomobject]@{
      vCenter=$VcName;Host=$h.Name;Cluster=$h.Parent.Name;DramGiB=[math]::Round($h.MemoryTotalGB,1)
      ActiveAvgGiB=$(if($act.Count){[math]::Round(($act|Measure-Object -Average).Average,1)})
      ActiveP95GiB=$(if($act.Count){[math]::Round((Get-Percentile $act .95),1)});ActiveMaxGiB=$(if($act.Count){[math]::Round(($act|Measure-Object -Maximum).Maximum,1)})
      TieringType=(Get-SafeProperty $h.ExtensionData.Hardware 'MemoryTieringType');NvmeTierPct=$null;Samples=$act.Count
      RequestedStartUtc=$start.ToUniversalTime().ToString('o');FirstSampleUtc=$(if($result.First){$result.First.ToUniversalTime().ToString('o')})
      LastSampleUtc=$(if($result.Last){$result.Last.ToUniversalTime().ToString('o')});IntervalUsed=$result.Interval;RequestedMetric=$result.RequestedMetric;SelectedMetric=$result.SelectedMetric;AvailableMemoryMetrics=$result.AvailableMemoryMetrics;CollectionStatus=$result.Status;CollectionError=$result.Error
    }) $VcName $h.Name
    Flush-AllEvidence;Set-ObjectComplete 'HostMemory' "$VcName|$($h.Name)"
  }
  if(-not$SkipVmStats){
    $vms=Get-VM -Server $Server|Where-Object{$_.PowerState-eq'PoweredOn'-and-not$_.ExtensionData.Config.Template};if($MaxVMs-gt 0){$vms=$vms|Select-Object -First $MaxVMs}
    foreach($vm in $vms){
      $result=Get-MemoryStatSamples $vm $start $vm.Name;$act=@($result.Stats|ForEach-Object{$_.Value/1MB})
      Add-EvidenceRow pcli_mem_vm ([pscustomobject]@{
        vCenter=$VcName;VM=$vm.Name;Uuid=$vm.ExtensionData.Config.Uuid;ConfiguredGiB=[math]::Round($vm.MemoryGB,1)
        ActiveAvgGiB=$(if($act.Count){[math]::Round(($act|Measure-Object -Average).Average,2)});ActiveP95GiB=$(if($act.Count){[math]::Round((Get-Percentile $act .95),2)})
        ActiveMaxGiB=$(if($act.Count){[math]::Round(($act|Measure-Object -Maximum).Maximum,2)});SamplesUsed=$act.Count
        RequestedStartUtc=$start.ToUniversalTime().ToString('o');FirstSampleUtc=$(if($result.First){$result.First.ToUniversalTime().ToString('o')})
        LastSampleUtc=$(if($result.Last){$result.Last.ToUniversalTime().ToString('o')});IntervalUsed=$result.Interval;RequestedMetric=$result.RequestedMetric;SelectedMetric=$result.SelectedMetric;AvailableMemoryMetrics=$result.AvailableMemoryMetrics;CollectionStatus=$result.Status;CollectionError=$result.Error
      }) $VcName
    }
  }
}

function Invoke-ClusterEvidenceV41 {
  param($Server,[string]$VcName)
  $started=[datetime]::UtcNow
  try{
    foreach($c in Get-Cluster -Server $Server){
      $cx=$c.ExtensionData;$lm=Get-SafePath $cx @('ConfigurationEx','LifecycleConfig','ManagedBy') ''
      $desired=Get-SafePath $cx @('ConfigurationEx','LifecycleConfig','DesiredSoftwareSpec','BaseImageSpec','Version') ''
      Add-EvidenceRow pcli_images ([pscustomobject]@{vCenter=$VcName;Cluster=$c.Name;LifecycleMode=$(if($lm){'Image'}else{'BaselineOrUnknown'});DesiredImageVersion=$desired;ComplianceStatus='NotExposed';VendorAddon='NotExposed';Components='NotExposed';LastComplianceCheckUtc='';HardwareSupportManager='NotExposed';CollectionStatus='Direct';CollectionError=''}) $VcName
      $vsc=Get-SafePath $cx @('ConfigurationEx','VsanConfigInfo') $null
      Add-EvidenceRow pcli_vsan ([pscustomobject]@{vCenter=$VcName;Cluster=$c.Name;Enabled=[bool](Get-SafeProperty $vsc 'Enabled' $false);Architecture=$(if(Get-SafeProperty $vsc 'VsanEsaEnabled' $false){'ESA'}elseif(Get-SafeProperty $vsc 'Enabled' $false){'OSA'}else{'Disabled'});Stretched='NotExposed';Witness='NotExposed';FaultDomains='NotExposed';Health='NotExposed';CapacityGiB=$null;FreeGiB=$null;CollectionStatus='Direct';CollectionError=''}) $VcName
    }
    Set-DatasetStatus pcli_images Produced '' $VcName @(Get-Cluster -Server $Server).Count (([datetime]::UtcNow-$started).TotalSeconds)
    Set-DatasetStatus pcli_vsan Produced '' $VcName @(Get-Cluster -Server $Server).Count (([datetime]::UtcNow-$started).TotalSeconds)
  }catch{Set-DatasetStatus pcli_images QueryFailed $_.Exception.Message $VcName 0 (([datetime]::UtcNow-$started).TotalSeconds);Set-DatasetStatus pcli_vsan QueryFailed $_.Exception.Message $VcName 0 (([datetime]::UtcNow-$started).TotalSeconds)}
  try{
    $si=Get-View -Server $Server ServiceInstance; $em=Get-View -Server $Server -Id $si.Content.ExtensionManager
    foreach($x in @($em.ExtensionList)){Add-EvidenceRow pcli_extensions ([pscustomobject]@{vCenter=$VcName;Name=$x.Key;Version=$x.Version;Company=$x.Company;Type=$x.Type;Status='Registered';CollectionStatus='Direct';CollectionError=''}) $VcName}
    Set-DatasetStatus pcli_extensions $(if(@($em.ExtensionList).Count){'Produced'}else{'SupportedEmpty'}) '' $VcName @($em.ExtensionList).Count 0
  }catch{Set-DatasetStatus pcli_extensions QueryFailed $_.Exception.Message $VcName}
}

function Register-EnvironmentDependentDatasetStatuses {
  param([string]$VcName)
  foreach($d in @('pcli_hsm','pcli_key_providers')){if(-not $script:DatasetTelemetry.Contains($d)){Set-DatasetStatus $d SupportedEmpty 'No provider objects returned by this lab vCenter.' $VcName 0 0}}
  foreach($d in @('pcli_vsandisks','pcli_nvme')){if(-not $script:DatasetTelemetry.Contains($d)){Set-DatasetStatus $d UnsupportedApi 'Detailed device inventory is not exposed by the active PowerCLI query path; boot and PCI evidence remain separate.' $VcName 0 0}}
  foreach($d in @('pcli_vibs','pcli_nicfw')){if(-not $script:DatasetTelemetry.Contains($d)){Set-DatasetStatus $d UnsupportedApi 'Host software or firmware query was not exposed through the lab ESXCLI path.' $VcName 0 0}}
}
function Write-EvidenceArchive {
  if($DisableEvidenceZip){return $null}
  $outFull=[IO.Path]::GetFullPath($script:OutDir).TrimEnd([IO.Path]::DirectorySeparatorChar,[IO.Path]::AltDirectorySeparatorChar)
  $parent=Split-Path -Parent $outFull
  if([string]::IsNullOrWhiteSpace($parent)){$parent=(Get-Location).Path}
  $zipPath=if($EvidenceZipPath){[IO.Path]::GetFullPath($EvidenceZipPath)}else{Join-Path $parent ("VCF-Evidence-{0}.zip"-f$script:RunId)}
  $zipFull=[IO.Path]::GetFullPath($zipPath)
  $insidePrefix=$outFull+[IO.Path]::DirectorySeparatorChar
  if($zipFull.StartsWith($insidePrefix,[StringComparison]::OrdinalIgnoreCase)-or$zipFull-eq$outFull){throw 'EvidenceZipPath must be outside OutDir to prevent recursive archive creation.'}
  $allowed=@('.csv','.json','.jsonl','.log','.txt')
  $excludeNames=@('collector-console.out.log','collector-console.err.log')
  $files=@(Get-ChildItem -LiteralPath $outFull -File|Where-Object{$_.Extension.ToLowerInvariant()-in$allowed-and$_.Name-notin$excludeNames-and$_.Extension-ne'.zip'})
  if($IncludeCollectorStateInZip){$files+=@(Get-ChildItem -LiteralPath $script:StateDir -File -Recurse|Where-Object{$_.Extension.ToLowerInvariant()-in$allowed})}
  $files=@($files|Sort-Object FullName -Unique)
  $fileRows=@($files|ForEach-Object{
    $relative=if($_.FullName.StartsWith($outFull,[StringComparison]::OrdinalIgnoreCase)){$_.FullName.Substring($outFull.Length).TrimStart('\\','/')}else{"collector-state/"+$_.Name}
    [pscustomobject]@{Path=$relative.Replace('\\','/');Bytes=$_.Length;Sha256=(Get-FileHash -Algorithm SHA256 -LiteralPath $_.FullName).Hash.ToLowerInvariant()}
  })
  $status=if($script:RunErrors.Count){'Partial'}elseif(@($script:DatasetStatus.Values|Where-Object{$_.Status-ne'Produced'}).Count){'CompleteWithGaps'}else{'Complete'}
  $packageManifest=Join-Path $outFull 'evidence-package-manifest.json'
  $manifest=[ordered]@{SchemaVersion='1.1';RunId=$script:RunId;CollectorVersion=$script:CollectorVersion;CreatedUtc=[datetime]::UtcNow.ToString('o');Status=$status;VCenters=@($script:VCenterNames);IncludesCollectorState=[bool]$IncludeCollectorStateInZip;Datasets=$script:DatasetStatus;Files=$fileRows}
  Write-AtomicText $packageManifest ($manifest|ConvertTo-Json -Depth 12)
  $files=@($files+(Get-Item -LiteralPath $packageManifest)|Sort-Object FullName -Unique)
  $zipDir=Split-Path -Parent $zipFull;if(-not(Test-Path -LiteralPath $zipDir)){New-Item -ItemType Directory -Force -Path $zipDir|Out-Null}
  Add-Type -AssemblyName System.IO.Compression;Add-Type -AssemblyName System.IO.Compression.FileSystem
  if(Test-Path $zipFull){Remove-Item -LiteralPath $zipFull -Force}
  $fs=[IO.File]::Open($zipFull,[IO.FileMode]::CreateNew);$za=[IO.Compression.ZipArchive]::new($fs,[IO.Compression.ZipArchiveMode]::Create,$false)
  try{
    foreach($file in $files){
      $rel=if($file.FullName.StartsWith($outFull,[StringComparison]::OrdinalIgnoreCase)){$file.FullName.Substring($outFull.Length).TrimStart('\\','/').Replace('\\','/')}else{"collector-state/"+$file.Name}
      $entry=$za.CreateEntry($rel,[IO.Compression.CompressionLevel]::Optimal);$input=$file.OpenRead();$output=$entry.Open();try{$input.CopyTo($output)}finally{$output.Dispose();$input.Dispose()}
    }
  }finally{$za.Dispose();$fs.Dispose()}
  $hash=(Get-FileHash -Algorithm SHA256 -LiteralPath $zipFull).Hash.ToLowerInvariant()
  $shaPath="$zipFull.sha256";[IO.File]::WriteAllText($shaPath,"$hash  $([IO.Path]::GetFileName($zipFull))"+[Environment]::NewLine,[Text.UTF8Encoding]::new($false))
  $result=[ordered]@{RunId=$script:RunId;Archive=$zipFull;Bytes=(Get-Item $zipFull).Length;Sha256=$hash;ChecksumFile=$shaPath;CreatedUtc=[datetime]::UtcNow.ToString('o');Status=$status}
  Write-AtomicText (Join-Path $outFull 'latest-evidence-archive.json') ($result|ConvertTo-Json -Depth 6)
  return $result
}
function Show-CollectionPlan {
  $names=Get-ActiveDatasetNames
  [pscustomobject]@{CollectorVersion=$script:CollectorVersion;Profile=$CollectionProfile;VCenterList=$VCenterList;VCenter=($script:VCenterNames-join', ');SharedCredential=[bool]$SharedCredential;ConnectionFailureAction=$ConnectionFailureAction;OutDir=$script:OutDir;Datasets=($names-join', ');IncludePci=[bool]$IncludePci;Days=$Days;VmFeatureSampleLimit=$VmFeatureSampleLimit}|Format-List
}

New-Item -ItemType Directory -Force -Path $script:OutDir|Out-Null
$script:VCenterNames=@(Get-VCenterNames $VCenterList)
if(-not$DryRun){Initialize-DurableState}
if($DryRun){Show-CollectionPlan;exit 0}
Assert-Prerequisites
Write-CollectorLog INFO "Acquiring credentials for $($script:VCenterNames.Count) vCenter(s) before connection or collection"
Initialize-Credentials $script:VCenterNames
Write-CollectorLog INFO 'All credentials acquired; starting connection preflight'
Invoke-ConnectionPreflight $script:VCenterNames
Write-CollectorLog INFO "Starting unified collector profile=$CollectionProfile for $($script:PreflightServers.Count) preflight-ready vCenter(s)"
foreach($vcName in @($script:VCenterNames|Where-Object{$script:PreflightServers.ContainsKey($_)})){
  if(Test-ObjectComplete 'VCenter' $vcName){Write-CollectorLog INFO 'Resume: vCenter already complete; skipping' $vcName;continue}
  $server=$script:PreflightServers[$vcName]
  try{
    Write-CollectorLog INFO 'Using preflight connection' $vcName
    switch($CollectionProfile){
      'VmFeaturesOnly'{Invoke-VmEvidence $server $vcName}
      'MemoryOnly'{Invoke-MemoryEvidence $server $vcName}
      default{Invoke-CoreInventory $server $vcName;Invoke-ClusterEvidenceV41 $server $vcName;Invoke-VmEvidence $server $vcName;if($CollectionProfile-in@('Full','Resume')){Invoke-HostDeepDive $server $vcName;Register-EnvironmentDependentDatasetStatuses $vcName;Invoke-MemoryEvidence $server $vcName}}
    }
    Flush-AllEvidence;Set-ObjectComplete 'VCenter' $vcName;Save-Checkpoint 'VCenterComplete' $vcName;Write-CollectorLog INFO 'Completed' $vcName
  }catch{$script:RunErrors.Add("${vcName}: $($_.Exception.Message)");Write-CollectorLog ERROR $_.Exception.Message $vcName;Flush-AllEvidence;Save-Checkpoint 'VCenterError' $vcName}
  finally{if($server){Disconnect-VIServer $server -Confirm:$false -ErrorAction SilentlyContinue;$script:PreflightServers.Remove($vcName)}}
}
Flush-AllEvidence
Complete-RequestedDatasets
$hasGaps=@($script:DatasetStatus.Values|Where-Object{$_.Status -ne 'Produced'}).Count -gt 0
$finalStatus=if($script:RunErrors.Count){'CompletedWithErrors'}elseif($hasGaps){'CompletedWithGaps'}else{'Completed'}
Write-RunManifest $finalStatus ($script:RunErrors -join ' | ')
Write-RunSummary $finalStatus
Write-CollectorLog INFO ("Collection complete. Status={0} Elapsed={1}" -f $finalStatus,([datetime]::UtcNow-$script:RunStartedUtc))
try{$archive=Write-EvidenceArchive;if($archive){Write-CollectorLog INFO ("Evidence archive created: {0} ({1} bytes) SHA256={2}"-f$archive.Archive,$archive.Bytes,$archive.Sha256)}}catch{Write-CollectorLog ERROR ("Evidence archive creation failed: {0}"-f$_.Exception.Message);$script:RunErrors.Add("Evidence archive: $($_.Exception.Message)")}
$ProgressPreference=$script:OriginalProgressPreference;$ConfirmPreference=$script:OriginalConfirmPreference;$InformationPreference=$script:OriginalInformationPreference
