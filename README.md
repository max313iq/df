# Azure Support Ticket Engine

PowerShell toolkit for bulk-creating Azure support tickets — primarily Batch quota increase requests. Two interfaces: CLI and WPF GUI.

## Prerequisites

- **PowerShell 7+** (`pwsh`).
- **Azure CLI** (`az`) — required for subscription discovery, Batch account enumeration, and token acquisition.
- A valid Azure subscription with permissions to create support tickets.

## Quick Start

### CLI

```powershell
# Dry run using Azure CLI token (no API calls)
./create_azure_support_tickets.ps1 -DryRun -TryAzCliToken $true

# Auto-discover Batch accounts and submit quota tickets
./create_azure_support_tickets.ps1 -AutoDiscoverRequests -TryAzCliToken $true

# Use an explicit bearer token
./create_azure_support_tickets.ps1 -Token $env:AZURE_BEARER_TOKEN
```

### GUI (Windows only)

```powershell
./create_azure_support_tickets_gui.ps1
```

Requires .NET Framework WPF assemblies. Discover accounts, configure settings, and execute from the graphical interface.

## Project Structure

```
├── create_azure_support_tickets.ps1       # CLI entry point
├── create_azure_support_tickets_gui.ps1   # WPF GUI entry point (Windows)
├── Test-ParityVerification.ps1            # 102 parity/unit tests
├── config/
│   ├── default-ticket-template.json       # Default ticket template
│   ├── ticket-template.schema.json        # JSON Schema for templates
│   └── azure-ticket-gui-profile.json      # Saved GUI profile
├── Modules/
│   ├── AzureSupport.TicketEngine.psm1     # Shared module (public API)
│   └── Private/
│       └── AzureSupport.TicketEngine.Core.ps1  # Core engine
└── AGENTS.md                              # Dev environment instructions
```

## Architecture

### Module (`AzureSupport.TicketEngine`)

The shared module provides a single API surface consumed by both CLI and GUI:

| Category | Functions |
|---|---|
| **Core Engine** | `Invoke-AzureSupportBatchQuotaRun`, `Invoke-AzureSupportBatchQuotaRunQueued` |
| **Discovery** | `Get-AzureSupportDiscoveryRows`, `Get-AzureRegionList`, `New-DiscoveryGridRow`, `Test-DiscoveryRegionValue` |
| **Templates** | `Get-TicketTemplate`, `Merge-TemplateDefaults`, `Resolve-EffectiveContactDetails`, `Resolve-EffectiveTemplateValues` |
| **Profiles** | `Get-RunProfile`, `Save-RunProfile`, `Convert-ProfileToUnifiedSchema`, `New-RunProfileSnapshot` |
| **Preflight** | `Test-AzureSupportPreFlight`, `ConvertTo-ValidatedRequestList` |
| **Validation** | `Test-NonEmptyString`, `Test-NumericRange`, `Test-EmailFormat`, `ConvertTo-TrimmedString`, `ConvertTo-EscapedString` |
| **Utilities** | `Convert-ToBoolValue`, `Convert-ToIntValue`, `Convert-ToStringArray`, `Get-ObjectMemberValue`, `Get-FirstDefinedValue` |

### Data Flow

```
CLI script → Import-Module → Invoke-AzureSupportTicketCli → Core.ps1 (script)
GUI script → Import-Module → Background jobs call module functions directly
```

### Profile Schema (v1)

Run profiles use a unified `profileVersion=1` schema with sections: `runSettings`, `execution`, `proxy`, `resume`, `defaults`, `ticket`, `ui`. Legacy flat profiles are automatically migrated on load.

## Configuration

### Ticket Template (`config/default-ticket-template.json`)

Defines contact details, default requests, quota parameters, and ticket metadata. Override individual fields via CLI parameters or GUI controls.

### Template Schema

See `config/ticket-template.schema.json` (JSON Schema draft 2020-12) for the full specification.

## Key Features

- **Template-driven**: Default contacts, requests, and ticket structure from JSON config.
- **Auto-discovery**: Enumerate Batch accounts across subscriptions via Azure CLI.
- **Multi-region**: One ticket per (account, region) pair with editable region lists.
- **Resume/retry**: Persisted run state enables continuing from interruption or retrying failures.
- **Proxy pools**: Rotate across multiple proxy endpoints per request.
- **Cancellation**: Graceful stop via signal file.
- **Dry run**: Validate everything without making API calls.
- **Artifact generation**: JSON/CSV result exports and redacted logs.

## Running Tests

```bash
TEMP=/tmp pwsh -File ./Test-ParityVerification.ps1
```

Runs 102 offline parity/unit tests covering module loading, template parsing, profile migration, region validation, preflight checks, and dry-run scenarios. No Azure credentials required.

## License

See repository root for license information.
