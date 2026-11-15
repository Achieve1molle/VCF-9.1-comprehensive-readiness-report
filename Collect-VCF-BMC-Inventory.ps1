[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$BmcList,
    [string]$OutDir = '.\vcfr-collect',
    [pscredential]$Credential = (Get-Credential -Message 'Read-only BMC account (iDRAC/iLO/CIMC)'),
    [ValidateRange(5,300)][int]$TimeoutSec = 30
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null

$pair = '{0}:{1}' -f $Credential.UserName, $Credential.GetNetworkCredential().Password
$headers = @{
    Authorization = 'Basic ' + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes($pair))
    Accept = 'application/json'
}
$useSkipCertificateCheck = $PSVersionTable.PSVersion.Major -ge 7
if (-not $useSkipCertificateCheck) {
    Add-Type @'
using System.Net;
using System.Security.Cryptography.X509Certificates;
public class TrustAllVcfr : ICertificatePolicy {
    public bool CheckValidationResult(ServicePoint s, X509Certificate c, WebRequest r, int p) { return true; }
}
'@
    [Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllVcfr
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
}

$systems = [System.Collections.Generic.List[object]]::new()
$firmware = [System.Collections.Generic.List[object]]::new()
$pcie = [System.Collections.Generic.List[object]]::new()
$storage = [System.Collections.Generic.List[object]]::new()
$diagnostics = [System.Collections.Generic.List[object]]::new()

function Add-Diagnostic {
    param([string]$Address,[string]$Category,[string]$Endpoint,[string]$Status,[int]$Rows,[string]$Detail)
    $diagnostics.Add([pscustomobject]@{
        BmcAddress=$Address; Category=$Category; Endpoint=$Endpoint; Status=$Status; Rows=$Rows; Detail=$Detail
    })
}

function Resolve-RedfishUri {
    param([string]$Base,[object]$Reference)
    if ($null -eq $Reference) { return $null }
    $path = if ($Reference -is [string]) { $Reference } elseif ($Reference.PSObject.Properties.Name -contains '@odata.id') { $Reference.'@odata.id' } else { $null }
    if ([string]::IsNullOrWhiteSpace([string]$path)) { return $null }
    if ($path -match '^https?://') { return [string]$path }
    return $Base.TrimEnd('/') + '/' + ([string]$path).TrimStart('/')
}

function Invoke-RedfishGet {
    param([string]$Base,[object]$Reference,[string]$Address,[string]$Category,[switch]$Quiet)
    $uri = Resolve-RedfishUri -Base $Base -Reference $Reference
    if (-not $uri) { return $null }
    try {
        $params = @{ Uri=$uri; Headers=$headers; Method='Get'; TimeoutSec=$TimeoutSec; ErrorAction='Stop' }
        if ($useSkipCertificateCheck) { $params.SkipCertificateCheck = $true }
        return Invoke-RestMethod @params
    }
    catch {
        if (-not $Quiet) {
            $code = ''
            try { $code = [int]$_.Exception.Response.StatusCode } catch {}
            Add-Diagnostic $Address $Category $uri 'RequestFailed' 0 ("HTTP {0}: {1}" -f $code,$_.Exception.Message)
        }
        return $null
    }
}

function Get-RedfishItems {
    param([string]$Base,[object]$Reference,[string]$Address,[string]$Category)
    $result = [System.Collections.Generic.List[object]]::new()
    if ($null -eq $Reference) { return @() }
    $refs = @()
    if ($Reference -is [System.Array]) { $refs = @($Reference) }
    elseif ($Reference -isnot [string] -and $Reference.PSObject.Properties.Name -contains 'Members') { $refs = @($Reference.Members) }
    else { $refs = @($Reference) }

    foreach ($ref in $refs) {
        $obj = if ($ref -is [string] -or ($ref.PSObject.Properties.Name -contains '@odata.id')) {
            Invoke-RedfishGet $Base $ref $Address $Category
        } else { $ref }
        if ($null -eq $obj) { continue }
        if ($obj.PSObject.Properties.Name -contains 'Members' -and $obj.Members) {
            foreach ($member in @($obj.Members)) {
                $item = Invoke-RedfishGet $Base $member $Address $Category
                if ($item) { $result.Add($item) }
            }
        } else { $result.Add($obj) }
    }
    return @($result)
}

function Get-PropertyValue {
    param([object]$Object,[string[]]$Names)
    if ($null -eq $Object) { return $null }
    foreach ($name in $Names) {
        if ($Object.PSObject.Properties.Name -contains $name) {
            $value = $Object.$name
            if ($null -ne $value -and "${value}" -ne '') { return $value }
        }
    }
    return $null
}

function Get-RelatedPcieDevices {
    param([string]$Base,[object]$System,[object[]]$Chassis,[string]$Address)
    $items = [System.Collections.Generic.List[object]]::new()
    $candidates = [System.Collections.Generic.List[object]]::new()
    if ($System.PSObject.Properties.Name -contains 'PCIeDevices') { $candidates.Add($System.PCIeDevices) }
    if ($System.Links -and $System.Links.PSObject.Properties.Name -contains 'PCIeDevices') { $candidates.Add($System.Links.PCIeDevices) }
    foreach ($ch in @($Chassis)) {
        if ($ch.PSObject.Properties.Name -contains 'PCIeDevices') { $candidates.Add($ch.PCIeDevices) }
        if ($ch.Links -and $ch.Links.PSObject.Properties.Name -contains 'PCIeDevices') { $candidates.Add($ch.Links.PCIeDevices) }
    }
    foreach ($candidate in $candidates) {
        foreach ($item in @(Get-RedfishItems $Base $candidate $Address 'PCIe')) { $items.Add($item) }
    }
    $seen = @{}
    return @($items | Where-Object {
        $key = [string](Get-PropertyValue $_ @('@odata.id','Id','Name'))
        if (-not $key) { $key = [guid]::NewGuid().ToString() }
        if ($seen.ContainsKey($key)) { $false } else { $seen[$key]=$true; $true }
    })
}

function Export-StableCsv {
    param([object[]]$Rows,[string[]]$Columns,[string]$Path)
    if (@($Rows).Count -gt 0) { $Rows | Select-Object $Columns | Export-Csv $Path -NoTypeInformation -Encoding UTF8 }
    else { Set-Content -Path $Path -Value (($Columns | ForEach-Object { '"' + ($_ -replace '"','""') + '"' }) -join ',') -Encoding UTF8 }
}

$addresses = @(Get-Content $BmcList | ForEach-Object { ($_ -split ',')[0].Trim() } | Where-Object { $_ -and $_ -notmatch '^\s*#' } | Select-Object -Unique)
foreach ($address in $addresses) {
    $base = 'https://' + $address
    Write-Host "== $address" -ForegroundColor Cyan
    $root = Invoke-RedfishGet $base '/redfish/v1' $address 'Root'
    if (-not $root) {
        $systems.Add([pscustomobject]@{ BmcAddress=$address; Error='Redfish root unreachable or authentication failed' })
        continue
    }

    $systemItems = @(Get-RedfishItems $base $root.Systems $address 'Systems')
    if ($systemItems.Count -eq 0) { Add-Diagnostic $address 'Systems' '/redfish/v1/Systems' 'NoRows' 0 'No ComputerSystem resource was returned.'; continue }
    $system = $systemItems | Select-Object -First 1
    $manager = @(Get-RedfishItems $base $root.Managers $address 'Managers') | Select-Object -First 1
    $chassis = @(Get-RedfishItems $base $root.Chassis $address 'Chassis')
    $bios = Invoke-RedfishGet $base ($system.'@odata.id' + '/Bios') $address 'BIOS' -Quiet
    $secureBoot = Invoke-RedfishGet $base ($system.'@odata.id' + '/SecureBoot') $address 'SecureBoot' -Quiet
    $attributes = if ($bios) { $bios.Attributes } else { $null }
    $trustedModule = @($system.TrustedModules) | Select-Object -First 1
    $processors = @(Get-RedfishItems $base $system.Processors $address 'Processors')
    $processor = $processors | Select-Object -First 1
    $memory = @(Get-RedfishItems $base $system.Memory $address 'Memory')
    $dimms = @($memory | Where-Object { [double](Get-PropertyValue $_ @('CapacityMiB')) -gt 0 })
    $serviceTag = if ($system.SKU) { $system.SKU } else { $system.SerialNumber }
    $bootMode = Get-PropertyValue $attributes @('BootMode')
    if (-not $bootMode -and $system.Boot) { $bootMode = $system.Boot.BootSourceOverrideMode }

    $systems.Add([pscustomobject]@{
        BmcAddress=$address; Vendor=$system.Manufacturer; Model=$system.Model; ServiceTag=$serviceTag; Serial=$system.SerialNumber
        HostName=$system.HostName; PowerState=$system.PowerState; Health=$system.Status.Health; BiosVersion=$system.BiosVersion
        BmcModel=$manager.Model; BmcFirmware=$manager.FirmwareVersion; BootMode=$bootMode
        SecureBoot=if ($secureBoot) { if ($secureBoot.SecureBootEnable) {'Enabled'} else {'Disabled'} } else {''}
        TpmVersion=$trustedModule.InterfaceType; TpmFirmware=$trustedModule.FirmwareVersion; TpmState=$trustedModule.Status.State
        TpmSetting=if ($attributes.TpmSecurity) {$attributes.TpmSecurity} elseif ($attributes.TpmState) {$attributes.TpmState} else {''}
        CpuModel=$processor.Model; CpuSockets=$processors.Count; CpuCoresPerSocket=$processor.TotalCores
        CpuFamily=$processor.ProcessorId.EffectiveFamily; CpuModelId=$processor.ProcessorId.EffectiveModel; CpuStepping=$processor.ProcessorId.Step
        DimmCount=$dimms.Count; MemoryGiB=[math]::Round((($dimms | Measure-Object CapacityMiB -Sum).Sum / 1024),0)
        DimmUnhealthy=@($dimms | Where-Object { $_.Status.Health -and $_.Status.Health -ne 'OK' }).Count
        DimmTypes=(($dimms | ForEach-Object { "$($_.MemoryDeviceType)@$($_.OperatingSpeedMhz)" } | Sort-Object -Unique) -join ';')
        Error=''
    })

    $updateService = if ($root.UpdateService) { Invoke-RedfishGet $base $root.UpdateService $address 'UpdateService' -Quiet } else { $null }
    $firmwareRef = if ($updateService -and $updateService.FirmwareInventory) { $updateService.FirmwareInventory } else { '/redfish/v1/UpdateService/FirmwareInventory' }
    $fwItems = @(Get-RedfishItems $base $firmwareRef $address 'Firmware')
    foreach ($item in $fwItems) {
        $firmware.Add([pscustomobject]@{BmcAddress=$address;ServiceTag=$serviceTag;Name=$item.Name;Id=$item.Id;Version=$item.Version;Updateable=$item.Updateable;Health=$item.Status.Health})
    }
    Add-Diagnostic $address 'Firmware' (Resolve-RedfishUri $base $firmwareRef) $(if($fwItems.Count){'Collected'}else{'NoRows'}) $fwItems.Count $(if($fwItems.Count){''}else{'FirmwareInventory returned no members or is not exposed.'})

    $pcieItems = @(Get-RelatedPcieDevices $base $system $chassis $address)
    foreach ($device in $pcieItems) {
        $fnRef = if ($device.PCIeFunctions) { $device.PCIeFunctions } elseif ($device.Links -and $device.Links.PCIeFunctions) { $device.Links.PCIeFunctions } else { $null }
        $functions = @(Get-RedfishItems $base $fnRef $address 'PCIeFunctions')
        if ($functions.Count -eq 0) { $functions = @($null) }
        foreach ($function in $functions) {
            $pcie.Add([pscustomobject]@{
                BmcAddress=$address;ServiceTag=$serviceTag;Device=$device.Name;Manufacturer=$device.Manufacturer;Model=$device.Model
                DeviceFirmware=$device.FirmwareVersion;VID=Get-PropertyValue $function @('VendorId','VendorID')
                DID=Get-PropertyValue $function @('DeviceId','DeviceID');SVID=Get-PropertyValue $function @('SubsystemVendorId','SubsystemVendorID')
                SSID=Get-PropertyValue $function @('SubsystemId','SubsystemID');Class=Get-PropertyValue $function @('DeviceClass','ClassCode')
                Slot=Get-PropertyValue $device @('Slot','Id');DeviceOdataId=$device.'@odata.id';FunctionOdataId=if($function){$function.'@odata.id'}else{''}
            })
        }
    }
    Add-Diagnostic $address 'PCIe' 'System/Chassis PCIeDevices' $(if($pcieItems.Count){'Collected'}else{'NoRows'}) @($pcie | Where-Object BmcAddress -eq $address).Count $(if($pcieItems.Count){''}else{'No standard Redfish PCIeDevices links were returned. Check BMC firmware, privileges, and vendor Redfish support.'})

    $storageItems = @(Get-RedfishItems $base $system.Storage $address 'Storage')
    foreach ($st in $storageItems) {
        $controllers = @($st.StorageControllers)
        if ($controllers.Count -eq 0) { $controllers = @($null) }
        foreach ($controller in $controllers) {
            $mode=''
            try { $mode=$st.Oem.Dell.DellController.ControllerMode } catch {}
            $storage.Add([pscustomobject]@{BmcAddress=$address;ServiceTag=$serviceTag;Kind='Controller';Name=$st.Name;Model=Get-PropertyValue $controller @('Model','Name');Firmware=Get-PropertyValue $controller @('FirmwareVersion');Mode=$mode;Protocol='';Media='';SizeGB='';LifeLeftPct='';Health=if($st.Status){Get-PropertyValue $st.Status @('Health')}else{''};Serial='';PartNumber='';OdataId=$st.'@odata.id'})
        }
        foreach ($driveRef in @($st.Drives)) {
            $drive = Invoke-RedfishGet $base $driveRef $address 'StorageDrive'
            if (-not $drive) { continue }
            $capacity = if ($drive.CapacityBytes) { [math]::Round(([double]$drive.CapacityBytes / 1GB),0) } else { '' }
            $life = if ($null -ne $drive.PredictedMediaLifeLeftPercent) { $drive.PredictedMediaLifeLeftPercent } else { '' }
            $storage.Add([pscustomobject]@{BmcAddress=$address;ServiceTag=$serviceTag;Kind='Drive';Name=Get-PropertyValue $drive @('Name','Id');Model=Get-PropertyValue $drive @('Model','Name');Firmware=Get-PropertyValue $drive @('Revision','FirmwareVersion');Mode='';Protocol=Get-PropertyValue $drive @('Protocol');Media=Get-PropertyValue $drive @('MediaType');SizeGB=$capacity;LifeLeftPct=$life;Health=if($drive.Status){Get-PropertyValue $drive.Status @('Health')}else{''};Serial=Get-PropertyValue $drive @('SerialNumber');PartNumber=Get-PropertyValue $drive @('PartNumber');OdataId=Get-PropertyValue $drive @('@odata.id')})
        }
    }
    $simpleItems = @(Get-RedfishItems $base $system.SimpleStorage $address 'SimpleStorage')
    foreach ($simple in $simpleItems) {
        foreach ($device in @($simple.Devices)) {
            $storage.Add([pscustomobject]@{BmcAddress=$address;ServiceTag=$serviceTag;Kind='SimpleStorage';Name=Get-PropertyValue $device @('Name','Id');Model=Get-PropertyValue $device @('Model');Firmware='';Mode='';Protocol=Get-PropertyValue $device @('Protocol');Media='';SizeGB='';LifeLeftPct='';Health=if($device.Status){Get-PropertyValue $device.Status @('Health')}else{''};Serial='';PartNumber='';OdataId=Get-PropertyValue $simple @('@odata.id')})
        }
    }
    $hostStorageCount=@($storage | Where-Object BmcAddress -eq $address).Count
    Add-Diagnostic $address 'Storage' 'ComputerSystem Storage/SimpleStorage' $(if($hostStorageCount){'Collected'}else{'NoRows'}) $hostStorageCount $(if($hostStorageCount){''}else{'No Storage or SimpleStorage rows were exposed. Check BMC firmware, privileges, controller mode, and Redfish implementation.'})
}

Export-StableCsv @($systems) @('BmcAddress','Vendor','Model','ServiceTag','Serial','HostName','PowerState','Health','BiosVersion','BmcModel','BmcFirmware','BootMode','SecureBoot','TpmVersion','TpmFirmware','TpmState','TpmSetting','CpuModel','CpuSockets','CpuCoresPerSocket','CpuFamily','CpuModelId','CpuStepping','DimmCount','MemoryGiB','DimmUnhealthy','DimmTypes','Error') (Join-Path $OutDir 'bmc_systems.csv')
Export-StableCsv @($firmware) @('BmcAddress','ServiceTag','Name','Id','Version','Updateable','Health') (Join-Path $OutDir 'bmc_firmware.csv')
Export-StableCsv @($pcie) @('BmcAddress','ServiceTag','Device','Manufacturer','Model','DeviceFirmware','VID','DID','SVID','SSID','Class','Slot','DeviceOdataId','FunctionOdataId') (Join-Path $OutDir 'bmc_pcie.csv')
Export-StableCsv @($storage) @('BmcAddress','ServiceTag','Kind','Name','Model','Firmware','Mode','Protocol','Media','SizeGB','LifeLeftPct','Health','Serial','PartNumber','OdataId') (Join-Path $OutDir 'bmc_storage.csv')
Export-StableCsv @($diagnostics) @('BmcAddress','Category','Endpoint','Status','Rows','Detail') (Join-Path $OutDir 'bmc_diagnostics.csv')

Write-Host ("Done: {0} systems, {1} firmware, {2} PCIe, {3} storage, {4} diagnostics -> {5}" -f $systems.Count,$firmware.Count,$pcie.Count,$storage.Count,$diagnostics.Count,$OutDir) -ForegroundColor Green
