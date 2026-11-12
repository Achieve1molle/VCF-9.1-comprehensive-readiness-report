# Comprehensive VCF 9.1 Readiness Report

A Presidio-branded assessment workflow for evaluating a VMware vSphere 8 estate against VMware Cloud Foundation 9.1 and ESXi 9.1 readiness requirements.

The solution combines RVTools inventory, read-only PowerCLI collectors, BMC/Redfish hardware evidence, current Broadcom knowledge articles, and Broadcom Compatibility Guide results. It produces a self-contained HTML report, an Excel evidence workbook, high-resolution charts, and a prioritized remediation register.

## Workflow

The solution follows the same five-stage workflow shown in the project graphic:

```text
1. CAPTURE       2. SANITIZE        3. GENERATE        4. DELIVER       5. RESTORE
Customer site    Optional            Presidio skill     Final product    Optional

RVTools          Tokenize names      Load evidence      HTML report      Restore names
PowerCLI         Retain technical    Review Broadcom    Excel workbook   using KEY.csv
BMC/Redfish      evidence            Validate BCG       PNG charts
                                    Run QA gates        Evidence JSON
```

### Stage 1: Capture

Minimum input:

- One complete RVTools `.xlsx` export per vCenter

Recommended read-only evidence:

- Core vCenter, cluster, host, network, datastore, PCI, VM encryption, and VM feature collection
- Host DNS, NTP, Secure Boot, TPM, image, VIB, NIC firmware, HSM, and vSAN disk collection
- Historical host and VM active-memory collection
- BMC/Redfish system, firmware, PCIe, and storage collection

### Stage 2: Sanitize

Sanitization is optional and runs at the customer site.

The sanitizer tokenizes environment identity while preserving compatibility evidence.

Tokenized or removed:

- vCenter, host, VM, cluster, datacenter, network, datastore, and folder names
- IP and MAC addresses
- UUIDs, serial numbers, and service tags
- Credentials, license keys, and general free-text notes

Preserved:

- Hardware vendor and model
- CPU model and series
- Guest operating system
- BIOS, firmware, and driver versions
- PCI VID, DID, SVID, and SSID
- Protocol, media, TPM, Secure Boot, and other technical readiness evidence

`KEY.csv` remains at the customer site and must not be included in the upload package.

### Stage 3: Generate

The Presidio readiness skill:

1. Validates the input files and evidence coverage.
2. Reviews current Broadcom KB and TechDocs guidance for VCF 9.1.
3. Attempts live Broadcom Compatibility Guide validation.
4. Produces a portable sidecar plan when direct BCG access is blocked or incomplete.
5. Reconciles RVTools, PowerCLI, BMC, BCG, driver, firmware, encryption, memory, and health evidence.
6. Generates the report and evidence workbook.
7. Runs QA, count-integrity, consistency, sanitization, and release gates.

### Stage 4: Deliver

Primary deliverables:

- Self-contained HTML readiness report
- Excel evidence workbook
- High-resolution PNG charts
- Verification evidence JSON
- BCG query plan or sidecar results when applicable
- Prioritized action and remediation register

### Stage 5: Restore

When sanitized deliverables need their original environment identifiers, the customer uses the retained `KEY.csv` and the Rev2 unsanitizer at the customer site.

---

## Repository contents

Recommended cleaned package layout:

```text
VCF91-Readiness/
├── README.md
├── WIKI.md
├── RunCollection.md
├── Workflow-VCF91-Readiness.png
├── collectors/
│   ├── Collect-VCF-Readiness-Rev2.ps1
│   ├── Collect-VCF-Host-DeepDive-Rev2.ps1
│   ├── Collect-VCF-Memory-Active-Rev2.ps1
│   └── Collect-VCF-BMC-Inventory.ps1
├── sanitization/
│   ├── Invoke-VCF-Sanitize-Rev2.ps1
│   └── Invoke-VCF-Unsanitize-Rev2.ps1
├── sidecar/
│   └── BCG-Sidecar-Collector/
└── skill/
    └── presidio-comprehensive-vcf-9-1-readiness-report.zip
```

Older collector revisions, intermediate skill packages, development reports, and regression artifacts should be stored outside the distribution package.

---

## Requirements

### Collection workstation

- Windows PowerShell 7 or later
- VMware PowerCLI
- Network access to each vCenter over HTTPS
- Network access to BMC interfaces when Redfish collection is used
- Read-only accounts with sufficient inventory privileges

### Sanitization workstation

- PowerShell 7
- ImportExcel module for Excel sanitization and restoration
- Protected local storage for `KEY.csv`

### Report-generation environment

- Python 3
- pandas
- openpyxl
- matplotlib
- Access to current Broadcom documentation
- Broadcom Compatibility Guide access when available

---

## Quick start

### 1. Define the workspace and vCenters

```powershell
$Root = 'C:\Script\VCF91-Readiness'
$VCenter = @(
    'pod01vcsa01.corp.achieve-1.com',
    'pod01vcsa02.corp.achieve-1.com'
)

New-Item -ItemType Directory -Force -Path `
    $Root, `
    "$Root\RVTools", `
    "$Root\Readiness", `
    "$Root\DeepDive", `
    "$Root\Memory", `
    "$Root\BMC", `
    "$Root\Sanitized", `
    "$Root\ToUpload" | Out-Null

Set-ExecutionPolicy -Scope Process Bypass
```

### 2. Export RVTools

Create one complete RVTools `.xlsx` export per vCenter and place the files in:

```text
C:\Script\VCF91-Readiness\RVTools
```

### 3. Run the core readiness collector

```powershell
.\Collect-VCF-Readiness-Rev2.ps1 `
    -VCenter $VCenter `
    -OutDir "$Root\Readiness" `
    -IncludePci
```

This collector prompts separately for each vCenter when a shared `-Credential` is not supplied.

### 4. Run the host deep-dive collector

```powershell
.\Collect-VCF-Host-DeepDive-Rev2.ps1 `
    -VCenter $VCenter `
    -OutDir "$Root\DeepDive"
```

### 5. Run the memory and NVMe collector

```powershell
.\Collect-VCF-Memory-Active-Rev2.ps1 `
    -VCenter $VCenter `
    -Days 30 `
    -MaxVMsPerVCenter 50 `
    -OutDir "$Root\Memory"
```

### 6. Run the BMC/Redfish collector

Create `bmc.txt` with one iDRAC, iLO, or CIMC address per line, then run:

```powershell
.\Collect-VCF-BMC-Inventory.ps1 `
    -BmcList .\bmc.txt `
    -OutDir "$Root\BMC"
```

### 7. Review coverage and diagnostics

```powershell
Get-ChildItem "$Root\Readiness","$Root\DeepDive","$Root\Memory","$Root\BMC" `
    -File -Filter *.csv |
    Select-Object DirectoryName, Name, Length |
    Sort-Object DirectoryName, Name |
    Format-Table -AutoSize
```

Review these files before continuing:

```text
pcli_vm_coverage.csv
pcli_deepdive_diagnostics.csv
pcli_mem_coverage.csv
pcli_mem_diagnostics.csv
bmc_diagnostics.csv
```

### 8. Sanitize when required

```powershell
$InputPaths = @(
    "$Root\RVTools",
    "$Root\Readiness",
    "$Root\DeepDive",
    "$Root\Memory",
    "$Root\BMC"
)

.\Invoke-VCF-Sanitize-Rev2.ps1 `
    -Path $InputPaths `
    -KeyFile "$Root\KEY.csv" `
    -OutDir "$Root\Sanitized" `
    -Verify
```

### 9. Create the upload archive

```powershell
Copy-Item "$Root\Sanitized\*" "$Root\ToUpload" -Recurse -Force

Compress-Archive `
    -Path "$Root\ToUpload\*" `
    -DestinationPath "$Root\VCF91-Readiness-Evidence.zip" `
    -Force
```

If sanitization is not required, copy the RVTools and collector directories directly into `ToUpload` before creating the ZIP.

### 10. Restore identifiers when required

```powershell
.\Invoke-VCF-Unsanitize-Rev2.ps1 `
    -Path .\Sanitized_Estate_VCF_9.1_Readiness_Report.html, `
          .\Sanitized_Estate_VCF_9.1_Readiness_Data.xlsx `
    -KeyFile "$Root\KEY.csv" `
    -OutDir "$Root\Restored"
```

For the complete collection procedure, see `RunCollection.md`.

---

## Evidence collected

### Core readiness

- vCenter version and build
- Cluster HA, DRS, EVC, lifecycle, vSAN, and stretched evidence
- Host hardware, ESXi build, BIOS, TPM, connection state, and lockdown mode
- VMkernel services, IP mode, MTU, and service roles
- Distributed switch version, MTU, uplinks, and LACP
- Datastore type and capacity
- PCI identifiers and drivers
- Complete VM encryption inventory
- Deterministic detailed VM-feature sample with coverage reporting

### Host deep dive

- DNS and NTP configuration
- Secure Boot and TPM
- Lockdown mode and SSH
- Image profile and OEM add-on
- VIB inventory and acceptance levels
- NIC driver and firmware
- vSAN disks and disk groups
- Stage-level diagnostics

### Memory and NVMe

- Host and VM active-memory average, P95, and maximum
- Sample counts and query status
- Memory-tier configuration
- Local NVMe candidates
- Boot-device identification
- Coverage and diagnostics

### BMC and Redfish

- Server manufacturer and model
- BIOS and BMC firmware
- CPU and DIMM inventory
- Secure Boot and TPM
- PCIe models, firmware, and identifiers
- Storage controllers and physical drives
- Health and diagnostic results

---

## Report capabilities

The report includes:

- Executive decision dashboard
- Per-vCenter architecture and readiness cards
- Cluster readiness and recommended pathways
- Gate-based domain scoring
- Hardware and Broadcom compatibility evidence
- CPU, server, adapter, driver, firmware, storage, and NVMe findings
- Lifecycle and version readiness
- vSphere 8 health and conversion prechecks
- Security, TPM, Secure Boot, VM encryption, vTPM, and cTPM readiness
- Network, VMkernel, distributed-switch, and uplink findings
- vSAN and datastore readiness
- HCX migration readiness and per-VM constraints
- Capacity, N+1, utilization, licensing, right-sizing, and memory tiering
- Evidence-completeness matrix
- Prioritized remediation action register
- Detailed source-verification appendix

The source environment is assumed to be vSphere 8 without NSX, SDDC Manager, or VCF Operations. Those products are not treated as missing prerequisites or limitations.

---

## Readiness terminology

| Status | Meaning |
|---|---|
| Verified eligible | Exact evidence supports readiness for VCF 9.1 or ESXi 9.1 |
| Ready with conditions | Supported with a documented restriction, deprecated component, or required configuration condition |
| Verify | Exact BCG or configuration evidence remains incomplete |
| Remediation available | A specific change can make the object eligible |
| Blocked | Current server, CPU, lifecycle, capacity, or other authoritative evidence prevents readiness |
| Not collected | The collector evidence required for the check is absent |

Missing evidence is never treated as unsupported, and missing evidence cannot create a passing score.

---

## Broadcom Compatibility Guide workflow

The skill first attempts live BCG access.

- If live BCG succeeds, the report uses the live results.
- If BCG access is blocked or incomplete, the skill exports a portable sidecar plan.
- The sidecar uses the same query and matching modules as the live workflow.
- Imported sidecar results must pass plan, checksum, freshness, and evidence validation.

The workflow supports:

- Exact server-model matching
- Canonical HPE, Cisco, and Dell partner matching
- ESXi 9.1 release validation
- CPU-series correlation
- PCI VID/DID/SVID/SSID lookup
- Driver and firmware evidence association
- Conservative unresolved handling for storage records lacking exact BCG identifiers

---

## Safety and security

- All collectors are read-only.
- No collector writes vCenter or BMC passwords to the CSV outputs.
- Use least-privilege inventory accounts.
- Pilot each collector against one vCenter and a small BMC list first.
- Keep separate output directories for each collection run.
- Review diagnostic files before interpreting empty outputs.
- Keep `KEY.csv` at the customer site.
- Do not upload credentials, unrestricted logs, or stale intermediate evidence.
- Protect restored and unsanitized reports according to customer data-handling requirements.

---

## Related documentation

- `RunCollection.md`: command-by-command collection and packaging runbook
- `WIKI.md`: architecture, evidence model, reporting logic, troubleshooting, and operating procedures
- `Workflow-VCF91-Readiness.png`: visual five-stage workflow

---

## Current release status

The report skill has passed:

- Active Python regression tests
- HTML and Excel generation
- 34 report QA checks
- Count-integrity checks
- HTML-to-analysis consistency checks
- Sanitization verification
- Leak scanning
- Release gating

External collectors remain subject to live validation against representative target vCenters and BMC implementations.

