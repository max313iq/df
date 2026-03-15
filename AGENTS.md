# AGENTS.md

## Cursor Cloud specific instructions

### Project overview

Azure Support Ticket Engine — a PowerShell toolkit for bulk-creating Azure support tickets (primarily Batch quota increase requests). Two interfaces: CLI (`create_azure_support_tickets.ps1`) and WPF GUI (`create_azure_support_tickets_gui.ps1`, Windows-only). No package managers, build steps, or containers — pure PowerShell + JSON config.

### Runtime

- **PowerShell 7+** (`pwsh`) is the sole runtime. Install via Microsoft package repo if missing.
- The `TEMP` environment variable must be set on Linux (e.g. `export TEMP=/tmp`); the test harness uses `$env:TEMP` which is only set automatically on Windows.

### Running tests

```
TEMP=/tmp pwsh -File /workspace/Test-ParityVerification.ps1
```

This runs 102 parity/unit tests covering module loading, template parsing, profile migration, region validation, preflight checks, and dry-run scenarios. All tests run offline — no Azure credentials required.

### Running the CLI

The CLI (`create_azure_support_tickets.ps1`) requires Azure credentials (bearer token or Azure CLI login) and network access to `management.azure.com`. It cannot be exercised end-to-end without a real Azure subscription. Dry-run mode also requires Azure CLI for auto-discovery unless requests are provided via the template.

### Key gotchas

- The GUI script (`create_azure_support_tickets_gui.ps1`) requires WPF/.NET Framework and only works on Windows.
- Path separators: the codebase uses Windows-style backslashes in `Join-Path` calls (e.g. `'Modules\AzureSupport.TicketEngine.psm1'`), but PowerShell 7 on Linux handles these correctly.
- The module (`Modules/AzureSupport.TicketEngine.psm1`) dynamically parses and loads functions from `Modules/Private/AzureSupport.TicketEngine.Core.ps1` using AST parsing at import time.
