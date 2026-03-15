# Azure Support Ticket Engine

PowerShell toolkit for bulk-creating Azure support tickets (primarily Azure Batch quota requests) from either:

- CLI: `create_azure_support_tickets.ps1`
- GUI (Windows only): `create_azure_support_tickets_gui.ps1`

The codebase uses a shared engine module (`Modules/AzureSupport.TicketEngine.psm1`) so GUI and CLI execute through the same core APIs and validation logic.

## Runtime Requirements

- PowerShell 7+ (`pwsh`)
- Azure CLI (`az`) for discovery and token acquisition flows
- Linux/macOS users: set `TEMP` before running tests

Example:

- `export TEMP=/tmp`

## Repository Layout

- `create_azure_support_tickets.ps1` - CLI entry point
- `create_azure_support_tickets_gui.ps1` - WPF GUI entry point (Windows only)
- `Modules/AzureSupport.TicketEngine.psm1` - public module surface and shared helper APIs
- `Modules/Private/AzureSupport.TicketEngine.Core.ps1` - execution pipeline, queueing, persistence, retry logic
- `config/default-ticket-template.json` - default template/contact settings and optional default requests
- `config/azure-ticket-gui-profile.json` - persisted GUI profile
- `Test-ParityVerification.ps1` - offline parity/unit tests

## Quick Start (CLI)

1. Authenticate Azure CLI if not using bearer token:
   - `az login`
2. Run with template requests in dry-run:
   - `pwsh -File ./create_azure_support_tickets.ps1 -DryRun`
3. Run explicit request list (example from PowerShell):
   - build `@([pscustomobject]@{ sub='...'; account='...'; region='eastus'; limit=680; quotaType='LowPriority' })`
   - pass via `-Requests`

> Note: Live submission requires valid Azure credentials and network access to `management.azure.com`.

## Running the GUI

- Windows only (WPF dependency)
- Start with:
  - `pwsh -File ./create_azure_support_tickets_gui.ps1`
- GUI uses the same module APIs for:
  - discovery (`Get-AzureSupportDiscoveryRows`)
  - region validation (`Test-DiscoveryRegionValue` + `Get-AzureRegionList`)
  - execution (`Invoke-AzureSupportBatchQuotaRun`)

## Validation and Discovery APIs

Public shared APIs now include:

- Discovery helpers:
  - `Get-AzureSupportDiscoveryRows`
  - `New-DiscoveryGridRow`
  - `Test-DiscoveryRegionValue`
  - `Get-AzureRegionList`
- Input validation helpers:
  - `Convert-ToSanitizedString`
  - `Test-NonEmptyString`
  - `Test-NumericRange`
  - `Test-EmailFormat`
  - `Escape-SpecialCharacters`
- Profile migration:
  - `Convert-ProfileToUnifiedSchema`

## Execution Pipeline

`Invoke-AzureSupportBatchQuotaRun` is the canonical batch execution path. It provides:

- request normalization (one account+region per request)
- preflight validation
- retry/backoff and throttling handling
- resume/retry-failed support via persisted run state
- optional dry-run mode
- JSON/CSV result exports and state snapshots

## Testing

Run the offline parity suite:

- `TEMP=/tmp pwsh -File /workspace/Test-ParityVerification.ps1`

The suite validates module exports, template parsing, profile migration, discovery/request prep parity, and resume/retry scenarios.

## Profile Schema

Unified profile schema is versioned (`profileVersion: 1`):

- `runSettings`
- `execution`
- `proxy`
- `resume`
- `defaults`
- `ticket`
- `ui`

Legacy flat profiles are auto-migrated by `Convert-ProfileToUnifiedSchema`.