# AGENTS.md

## Cursor Cloud specific instructions

### Project overview

This is a pure **PowerShell** project (no package.json, no build system, no Docker). It provides CLI and GUI tools for automating Azure Batch quota-increase support ticket creation. See `README.md` for a high-level description.

### Key components

| Component | Path | Notes |
|---|---|---|
| CLI tool | `create_azure_support_tickets.ps1` | Cross-platform (PowerShell 7+) |
| GUI tool | `create_azure_support_tickets_gui.ps1` | **Windows-only** (WPF/XAML) — cannot run on Linux |
| Engine module | `Modules/AzureSupport.TicketEngine.psm1` | Shared logic for CLI and GUI |
| Engine core | `Modules/Private/AzureSupport.TicketEngine.Core.ps1` | Internal functions loaded by the module |
| Parity tests | `Test-ParityVerification.ps1` | Runs offline without Azure credentials |
| Config | `config/default-ticket-template.json` | Default ticket template and contact details |

### Running tests

The parity test suite runs without Azure credentials and validates module loading, utility functions, template parsing, profile migration, and dry-run scenarios:

```bash
TEMP=/tmp pwsh -File Test-ParityVerification.ps1
```

**Important**: On Linux, `$env:TEMP` is not set by default. You **must** set `TEMP=/tmp` (or any valid temp directory) when invoking any script that uses `$env:TEMP`, otherwise the script will fail with `Cannot bind argument to parameter 'Path' because it is null`.

### Expected test failures (environment-related)

Two tests fail on Linux without Azure CLI authentication — these are expected:

- `Get-AzureRegionList` — requires `$script:CachedAzureRegions` populated from Azure CLI
- `Invoke-AzCommand exported` — internal function, not exported from the module

### Running the CLI (dry-run)

To validate CLI functionality without Azure credentials:

```bash
TEMP=/tmp pwsh -Command '
  Import-Module ./Modules/AzureSupport.TicketEngine.psm1 -Force
  $template = Get-TicketTemplate -Path ./config/default-ticket-template.json
  $defaults = Merge-TemplateDefaults -Template $template
  # Build and validate requests from template
  $requests = @()
  foreach ($dr in $template.defaultRequests) {
      $requests += [pscustomobject]@{
          sub = [string]$dr.sub; account = [string]$dr.account; region = [string]$dr.region
          limit = [int]$defaults.NewLimit; quotaType = [string]$defaults.QuotaType
      }
  }
  Test-AzureSupportPreFlight -Requests $requests -Token "test" -TryAzCliToken $false `
      -DelaySeconds ([int]$defaults.DelaySeconds) -MaxRetries ([int]$defaults.MaxRetries) `
      -BaseRetrySeconds ([int]$defaults.BaseRetrySeconds) -RequestsPerMinute ([int]$defaults.RequestsPerMinute)
'
```

### Notes

- No linter is configured for this project (pure PowerShell, no PSScriptAnalyzer config).
- No build step is required — scripts run directly via `pwsh`.
- End-to-end ticket creation requires Azure CLI (`az`) authenticated with subscriptions that have Azure Batch resources.
- The GUI (`create_azure_support_tickets_gui.ps1`) requires Windows with WPF/.NET Framework — it cannot run on Linux.
