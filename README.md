# Presidio Comprehensive VCF 9.1 Readiness Report

Version 26 provides an evidence-driven assessment of a VMware vSphere estate for VMware Cloud Foundation 9.1 and ESXi 9.1 readiness. The solution combines RVTools, read-only PowerCLI evidence, optional BMC/Redfish inventory, Broadcom guidance, and Broadcom Compatibility Guide evidence to produce an HTML report, Excel evidence workbook, PNG charts, verification records, and a prioritized remediation register.

## What changed in Version 26

- Adds the unified `Collect-VCF-Evidence.ps1` collection workflow.
- Retains `Collect-VCF-BMC-Inventory.ps1` for optional BMC/Redfish hardware evidence.
- Supports sanitization and restoration either through the Skill workflow or customer-site PowerShell.
- Documents the built-in BCG sidecar workflow used by Copilot Agent Store.
- Documents the external/local sidecar workflow used with AI Studio.
- Supports a pre-built BCG cache for repeatable or disconnected report generation.
- Preserves evidence precedence, diagnostic states, leak scanning, release gates, and exact PCI matching.

## Minimum and recommended inputs

### Minimum

- One complete RVTools `.xlsx` export from one vCenter, with all worksheets retained.

### Recommended

- One RVTools export per vCenter.
- `Collect-VCF-Evidence.ps1` output for vCenter, cluster, host, network, datastore, vSAN, lifecycle, security, VM encryption, feature, utilization, and diagnostic evidence.
- `Collect-VCF-BMC-Inventory.ps1` output for server, firmware, PCIe, storage, TPM, Secure Boot, and Redfish diagnostics.
- Live BCG access, the Copilot built-in sidecar, an AI Studio sidecar return package, or a validated pre-built BCG cache.

## Version 26 workflow

1. **Capture** RVTools, PowerCLI, and optional BMC/Redfish evidence at the customer site.
2. **Sanitize** when required, using the Skill or `Invoke-VCF-Sanitize-Rev2.ps1`.
3. **Generate** the readiness assessment in Copilot or AI Studio.
4. **Deliver** HTML, Excel, PNG, JSON verification, and prioritized actions.
5. **Restore** customer identifiers when required, using the Skill or `Invoke-VCF-Unsanitize-Rev2.ps1` with the customer-retained `KEY.csv`.

## Collection

### Unified PowerCLI collection

```powershell
.\Collect-VCF-Evidence.ps1 `
  -VCenterList .\vcenters.txt `
  -PasswordFile .\vcenter-credentials.csv `
  -AllowClearTextCredentials `
  -CollectionProfile Full `
  -OutputRoot .\Evidence
```

The unified collector is designed for unattended, long-running, multi-vCenter collection. Version 26 collection logic includes progress logging, stage checkpoints, host-level checkpoints, diagnostics, and recovery behavior for large estates. Use read-only vCenter accounts. Protect any credential file according to customer policy and remove it before packaging evidence.

### BMC/Redfish collection

```powershell
.\Collect-VCF-BMC-Inventory.ps1 `
  -BmcList .\bmc.txt `
  -OutDir .\Evidence\BMC
```

Supported targets include Dell iDRAC, HPE iLO, Cisco CIMC, and standards-compliant Redfish endpoints. Review `bmc_diagnostics.csv` whenever PCIe, storage, firmware, or physical-drive evidence is empty or incomplete.

## Sanitization and restoration

### Skill-managed workflow

Use the Skill-managed workflow when the Skill runtime has the required input and output access. Select sanitized processing before report generation. The Skill creates a token mapping and applies the same identity-removal rules used by the PowerShell workflow. Keep the resulting `KEY.csv` at the customer site and never upload it with sanitized evidence.

For restoration, provide the generated report and retained `KEY.csv` to the restoration workflow in the approved customer environment. Technical values remain unchanged; only mapped identifiers are restored.

### PowerShell workflow

Sanitize before upload:

```powershell
.\Invoke-VCF-Sanitize-Rev2.ps1 `
  -Path .\Evidence `
  -KeyFile .\KEY.csv `
  -OutDir .\Sanitized `
  -Verify
```

Restore approved deliverables:

```powershell
.\Invoke-VCF-Unsanitize-Rev2.ps1 `
  -Path .\Sanitized\VCF_Readiness_Report.html, .\Sanitized\VCF_Readiness_Data.xlsx `
  -KeyFile .\KEY.csv `
  -OutDir .\Restored
```

Sanitization tokenizes customer-identifying names, IP addresses, MAC addresses, UUIDs, serial numbers, service tags, domains, email addresses, credentials, license values, and free-text identity. Hardware models, CPU models, BIOS, firmware, drivers, PCI identifiers, protocols, operating-system versions, TPM, and Secure Boot evidence are preserved for compatibility analysis.

## Broadcom Compatibility Guide execution options

### Copilot Agent Store

The Copilot package uses the built-in sidecar workflow when direct BCG access is unavailable or incomplete. The Skill creates or consumes a validated query plan, applies checksum validation, imports returned BCG results, and runs the same evidence gates used for live lookup. Do not add unsupported `.html` fixtures or generated reports to `skill.zip`.

### AI Studio

AI Studio can use a local or externally hosted sidecar with custom tools/actions. The sidecar executes the BCG query plan in an authorized network context, returns machine-readable result JSON, and preserves query-plan and result checksums. AI Studio deployments can automate sidecar execution through Python, PowerShell, or an approved custom action.

### Pre-built BCG cache

A pre-built cache is recommended when:

- Internet access is restricted.
- The same server and component population is assessed repeatedly.
- Deterministic regression or repeatable customer reporting is required.
- Live BCG access is slow or unavailable.

The cache must include provenance, creation date, target ESXi/VCF release, normalized vendor/model keys, pagination completeness, query-plan checksum, result checksum, and source references. A cache entry is evidence only for the exact identifiers and release represented. Missing or failed cache lookups remain `Verify`; they are not `Unsupported`.

## Outputs

- Branded self-contained HTML readiness report
- Excel evidence and remediation workbook
- High-resolution PNG charts
- BCG verification JSON and source references
- Prioritized remediation actions and validation methods
- Evidence completeness, diagnostic, and release-gate results

## Security requirements

- All collectors are read-only.
- Credentials must not be written to evidence outputs.
- `KEY.csv` remains at the customer site.
- Never upload credentials, unrestricted logs, customer-retained keys, or stale evidence.
- Sanitized deliverables remain customer technical information.
- Unsanitized and restored deliverables require customer-approved handling controls.

## Upload package requirements

The Copilot upload ZIP must contain `SKILL.md` and its required runtime resources. Exclude generated `.html` reports, test fixtures such as `WI18_visual_fixture.html`, development outputs, temporary caches, customer evidence, credentials, and `KEY.csv`. Runtime JSON resources may remain when referenced by the Skill or scripts.

See `WIKI.md` for architecture, evidence semantics, troubleshooting, operating procedures, and deployment details.
