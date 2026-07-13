# Windows Custom Inventory for Azure Log Analytics

This PowerShell script collects detailed hardware and software inventory from Windows devices and sends the data to an Azure Log Analytics workspace. It was designed for use by **PowerStacks BI for Intune** customers to extend inventory visibility beyond what Intune natively provides.

## Overview

The script provides granular control over the type of inventory collected, including device details, installed applications, drivers, warranty information, and Microsoft 365 metadata. It supports modular configuration, compression, and safe ingestion into Log Analytics under size constraints.

For implementation guidance and integration with the BI for Intune reporting solution, refer to the documentation below.

🔗 [Windows Inventory Collection Script – PowerStacks BI for Intune Documentation](https://docs.powerstacks.com/bi-for-intune/installation/custom-inventory/windows-inventory-collection-script/)

## Features

- Application inventory (system and user scope)
- Device hardware inventory (CPU, memory, disks, monitors, chassis, etc.)
- Microsoft 365 versioning, channel, and update insights
- Driver inventory from both PnP and optional updates
- Warranty information lookup (Dell, Lenovo, Getac)
- Compressed and Base64-encoded payloads for Azure ingestion
- Compatible with Intune (SYSTEM context), GPO, or Task Scheduler

## Parameters

| Parameter                 | Description                                                                 |
|---------------------------|-----------------------------------------------------------------------------|
| `CustomerId`              | Log Analytics Workspace ID                                                  |
| `SharedKey`               | Primary Key for the workspace                                               |
| `CollectDeviceInventory`  | Enable or disable device inventory collection (default: `$true`)           |
| `CollectAppInventory`     | Enable or disable application inventory (default: `$true`)                 |
| `CollectDriverInventory`  | Enable or disable driver inventory (default: `$true`)                      |
| `RemoveBuiltInMonitors`   | Exclude internal monitors from results (default: `$false`)                 |
| `InventoryDateFormat`     | Timestamp formatting for final status output (default: `"MM-dd HH:mm"`)    |

## Ingestion & authentication

`$LogAPIMode` (top of the script) selects how data reaches Log Analytics:

- **`DataCollectorAPI`** – the legacy HTTP Data Collector API (workspace ID + shared key).
- **`LogIngestionAPI`** (recommended) – the modern Logs Ingestion API, posting to a Data Collection Endpoint (DCE) / Data Collection Rule (DCR).

Under `LogIngestionAPI`, the token used to post to the DCR can be obtained two ways:

1. **Client secret (classic).** An Entra app registration client id + secret embedded in the script (`$ClientId` / `$ClientSecret`). Simple, but every device carries a reusable secret.
2. **PowerStacks Entra Token Broker (secretless, recommended).** When `$BrokerUrl` is set, the device authenticates to the broker with its **Entra Join certificate** (mutual TLS), receives a short-lived signed assertion, and exchanges it for the inventory managed identity's token — **no secret on the device and no relay hop**. The upload to the DCR is otherwise identical.

### Broker settings

| Setting          | Description                                                                                                   |
|------------------|---------------------------------------------------------------------------------------------------------------|
| `$BrokerUrl`     | The Entra Token Broker URL, e.g. `https://<broker>.azurewebsites.net`. Leave the placeholder to disable broker mode. |
| `$BrokerClientId`| The broker's **inventory user-assigned managed identity** client id (broker deploy output `inventoryIdentityClientId`). Not the app registration. |

Precedence: if `$BrokerUrl` is set it is used; otherwise the script falls back to `$ClientId` / `$ClientSecret`. `$TenantId`, `$DceURI`, and `$DcrImmutableId` are required in both cases.

Requirements for broker mode:
- The device must be **Entra joined** (or hybrid-joined) and the script must run as **SYSTEM / elevated** — the Entra Join certificate's private key is TPM-bound and machine-scoped.
- The broker's inventory managed identity must hold **Monitoring Metrics Publisher** on the target DCR.

```powershell
# LogIngestionAPI via the broker (secretless)
$LogAPIMode     = "LogIngestionAPI"
$TenantId       = "<tenant guid>"
$DceURI         = "https://<dce>.<region>.ingest.monitor.azure.com"
$DcrImmutableId = "dcr-xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
$BrokerUrl      = "https://<broker>.azurewebsites.net"
$BrokerClientId = "<inventoryIdentityClientId>"
# leave $ClientId / $ClientSecret as placeholders
```

## Usage

```powershell
.\InventoryCollector.ps1 -CustomerId "<YourWorkspaceID>" -SharedKey "<YourPrimaryKey>"
```

### Optional Vendor Configuration

For warranty collection, set the appropriate vendor credentials in the script:

```powershell
$WarrantyDellClientID = "<your Dell API client ID>"
$WarrantyDellClientSecret = "<your Dell API secret>"
$WarrantyLenovoClientID = "<your Lenovo API key>"
```

## Output

Data is posted to custom tables in Log Analytics:

- `PowerStacksDeviceInventory`
- `PowerStacksAppInventory`
- `PowerStacksDriverInventory`

Payloads are compressed, encoded, and split into safe chunks to meet Azure ingestion limits.

## Requirements

- PowerShell 5.1 or later
- Azure Log Analytics Workspace
- API access for vendor warranty data (optional)
- Execution context with network access (SYSTEM context in Intune is supported)

## License

MIT License

This script is provided as-is without warranty. Test thoroughly before deploying in production.

---

This script is maintained by the PowerStacks team and intended for integration with [BI for Intune](https://powerstacks.com/bi-for-intune/), a reporting solution built for Microsoft Intune environments.
