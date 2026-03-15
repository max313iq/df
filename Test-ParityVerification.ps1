[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$engineModulePath = Join-Path $scriptRoot 'Modules\AzureSupport.TicketEngine.psm1'
$templatePath = Join-Path $scriptRoot 'config\default-ticket-template.json'
$guiProfilePath = Join-Path $scriptRoot 'config\azure-ticket-gui-profile.json'
$testResultDir = Join-Path $env:TEMP 'AzureSupportParityTest'

$pass = 0
$fail = 0
$tests = New-Object System.Collections.Generic.List[object]

function Write-TestResult {
    param([string]$Name, [bool]$Passed, [string]$Detail = '')
    $status = if ($Passed) { 'PASS' } else { 'FAIL' }
    $color = if ($Passed) { 'Green' } else { 'Red' }
    Write-Host "  [$status] $Name" -ForegroundColor $color
    if (-not [string]::IsNullOrWhiteSpace($Detail)) {
        Write-Host "         $Detail" -ForegroundColor Gray
    }
    $script:tests.Add([pscustomobject]@{ Name = $Name; Passed = $Passed; Detail = $Detail })
    if ($Passed) { $script:pass++ } else { $script:fail++ }
}

if (-not (Test-Path -LiteralPath $testResultDir)) {
    New-Item -ItemType Directory -Path $testResultDir -Force | Out-Null
}

# -------------------------------------------------------------------
Write-Host "`n=== Module Loading ===" -ForegroundColor Cyan
# -------------------------------------------------------------------

try {
    Import-Module -Name $engineModulePath -Force -ErrorAction Stop | Out-Null
    Write-TestResult -Name 'Engine module imports successfully' -Passed $true
}
catch {
    Write-TestResult -Name 'Engine module imports successfully' -Passed $false -Detail $_.Exception.Message
    Write-Host "`nCannot proceed without engine module." -ForegroundColor Red
    exit 1
}

$exportedCommands = (Get-Module AzureSupport.TicketEngine).ExportedCommands.Keys
$requiredExports = @(
    'Invoke-AzureSupportBatchQuotaRun',
    'Invoke-AzureSupportBatchQuotaRunQueued',
    'Invoke-AzureSupportTicketCli',
    'Get-AzureSupportDiscoveryRows',
    'Test-AzureSupportPreFlight',
    'Get-TicketTemplate',
    'Merge-TemplateDefaults',
    'Get-TicketTemplatePath',
    'Get-ObjectMemberValue',
    'Convert-ToBoolValue',
    'Convert-ToIntValue',
    'Convert-ToStringArray',
    'Get-FirstDefinedValue',
    'New-DiscoveryGridRow',
    'Test-DiscoveryRegionValue',
    'Split-DiscoveryFilterList',
    'ConvertTo-DiscoveryCollection',
    'Expand-AccountRegionRequests',
    'Resolve-RequestFieldValue',
    'New-RunProfileSnapshot',
    'Get-RunProfile',
    'Save-RunProfile',
    'Convert-ProfileToUnifiedSchema',
    'Get-AzureRegionList'
)

foreach ($fn in $requiredExports) {
    $found = $exportedCommands -contains $fn
    Write-TestResult -Name "Module exports '$fn'" -Passed $found
}

# -------------------------------------------------------------------
Write-Host "`n=== Template Loading ===" -ForegroundColor Cyan
# -------------------------------------------------------------------

try {
    $template = Get-TicketTemplate -Path $templatePath
    Write-TestResult -Name 'Template loads from default path' -Passed ($null -ne $template)
}
catch {
    Write-TestResult -Name 'Template loads from default path' -Passed $false -Detail $_.Exception.Message
}

try {
    $defaults = Merge-TemplateDefaults -Template $template
    $hasDefaults = ($null -ne $defaults -and $defaults.Count -gt 0)
    Write-TestResult -Name 'Merge-TemplateDefaults returns populated result' -Passed $hasDefaults
}
catch {
    Write-TestResult -Name 'Merge-TemplateDefaults returns populated result' -Passed $false -Detail $_.Exception.Message
}

$requiredDefaultKeys = @('DelaySeconds', 'RequestsPerMinute', 'MaxRetries', 'BaseRetrySeconds', 'NewLimit', 'QuotaType', 'Title', 'ContactFirstName', 'ContactLastName', 'ContactEmail')
foreach ($key in $requiredDefaultKeys) {
    $val = $defaults[$key]
    Write-TestResult -Name "Template default '$key' is populated" -Passed ($null -ne $val -and "$val" -ne '')
}

# -------------------------------------------------------------------
Write-Host "`n=== Region Discovery Validation ===" -ForegroundColor Cyan
# -------------------------------------------------------------------

$regionCheck1 = Test-DiscoveryRegionValue -Region '' -DiscoveredRegions @('eastus', 'westus2') -DefaultRegion 'eastus'
Write-TestResult -Name 'Empty region defaults to eastus' -Passed ($regionCheck1.Region -eq 'eastus')
Write-TestResult -Name 'Empty region produces warning' -Passed ($regionCheck1.Warnings.Count -gt 0)

$regionCheck2 = Test-DiscoveryRegionValue -Region 'westus2' -DiscoveredRegions @('eastus', 'westus2')
Write-TestResult -Name 'Known region passes validation' -Passed ($regionCheck2.IsValid -and $regionCheck2.Warnings.Count -eq 0)

$regionCheck3 = Test-DiscoveryRegionValue -Region 'unknownregion' -DiscoveredRegions @('eastus', 'westus2')
Write-TestResult -Name 'Unknown region produces warning' -Passed ($regionCheck3.Warnings.Count -gt 0)
Write-TestResult -Name 'Unknown region still valid (no error)' -Passed ($regionCheck3.IsValid)

$regionCheck4 = Test-DiscoveryRegionValue -Region 'eastus' -DiscoveredRegions @()
Write-TestResult -Name 'Region with no discovered list passes' -Passed ($regionCheck4.IsValid -and $regionCheck4.Warnings.Count -eq 0)

# -------------------------------------------------------------------
Write-Host "`n=== Azure Region List ===" -ForegroundColor Cyan
# -------------------------------------------------------------------

try {
    $regionList = @(Get-AzureRegionList)
    Write-TestResult -Name 'Get-AzureRegionList returns results' -Passed ($regionList.Count -gt 0) -Detail "Count: $($regionList.Count)"
    Write-TestResult -Name 'Region list contains eastus' -Passed ($regionList -contains 'eastus')
    Write-TestResult -Name 'Region list contains westeurope' -Passed ($regionList -contains 'westeurope')
    Write-TestResult -Name 'Region list is sorted' -Passed ($regionList[0] -le $regionList[-1])
}
catch {
    Write-TestResult -Name 'Get-AzureRegionList' -Passed $false -Detail $_.Exception.Message
}

# -------------------------------------------------------------------
Write-Host "`n=== Discovery Grid Row ===" -ForegroundColor Cyan
# -------------------------------------------------------------------

$gridRow = New-DiscoveryGridRow -Id 'test-1' -SubscriptionId 'sub-123' -TenantId 'tenant-abc' -AccountName 'myaccount' -Region 'westeurope' -DiscoveredRegions @('westeurope', 'eastus')
Write-TestResult -Name 'Grid row has correct Id' -Passed ($gridRow.Id -eq 'test-1')
Write-TestResult -Name 'Grid row has correct Region' -Passed ($gridRow.Region -eq 'westeurope')
Write-TestResult -Name 'Grid row default Selected is true' -Passed ($gridRow.Selected -eq $true)
Write-TestResult -Name 'Grid row Status is Discovered' -Passed ($gridRow.Status -eq 'Discovered')
Write-TestResult -Name 'Grid row DiscoveredRegions preserved' -Passed ($gridRow.DiscoveredRegions.Count -eq 2)

$gridRowEmpty = New-DiscoveryGridRow -Id 'test-2' -SubscriptionId 'sub-456' -TenantId '' -AccountName 'acct2' -Region ''
Write-TestResult -Name 'Grid row empty region defaults to eastus' -Passed ($gridRowEmpty.Region -eq 'eastus')

# -------------------------------------------------------------------
Write-Host "`n=== Invoke-AzCommand Parameter Fix ===" -ForegroundColor Cyan
# -------------------------------------------------------------------

$cmdInfo = Get-Command Invoke-AzCommand -ErrorAction SilentlyContinue
if ($null -ne $cmdInfo) {
    $paramNames = $cmdInfo.Parameters.Keys
    Write-TestResult -Name 'Invoke-AzCommand has CommandArgs parameter' -Passed ($paramNames -contains 'CommandArgs')
    Write-TestResult -Name 'Invoke-AzCommand does not shadow $Args' -Passed ($paramNames -notcontains 'Args')
} else {
    Write-TestResult -Name 'Invoke-AzCommand exported' -Passed $false -Detail 'Function not found'
}

# -------------------------------------------------------------------
Write-Host "`n=== Retry Parameter Alignment ===" -ForegroundColor Cyan
# -------------------------------------------------------------------

$preflight = Test-AzureSupportPreFlight `
    -Requests @([pscustomobject]@{ sub = 'sub-1'; account = 'acct-1'; region = 'eastus'; limit = 100; quotaType = 'LowPriority' }) `
    -Token 'test-token' -TryAzCliToken $false -DelaySeconds 5 -MaxRetries 3 -BaseRetrySeconds 10 -RequestsPerMinute 2
Write-TestResult -Name 'Preflight validation passes for valid request' -Passed $preflight.IsValid

$preflightDup = Test-AzureSupportPreFlight `
    -Requests @(
        [pscustomobject]@{ sub = 'sub-1'; account = 'acct-1'; region = 'eastus' },
        [pscustomobject]@{ sub = 'sub-1'; account = 'acct-1'; region = 'eastus' }
    ) `
    -Token 'test-token' -TryAzCliToken $false -DelaySeconds 5 -MaxRetries 3 -BaseRetrySeconds 10 -RequestsPerMinute 2
Write-TestResult -Name 'Preflight detects duplicate requests' -Passed (-not $preflightDup.IsValid)

# -------------------------------------------------------------------
Write-Host "`n=== Profile Schema Migration ===" -ForegroundColor Cyan
# -------------------------------------------------------------------

$flatProfile = [pscustomobject]@{
    DelaySeconds = 10; RequestsPerMinute = 5; MaxRetries = 3; BaseRetrySeconds = 15
    NewLimit = 500; QuotaType = 'Dedicated'; Title = 'Test Ticket'
    ContactFirstName = 'John'; ContactLastName = 'Doe'; ContactEmail = 'john@example.com'
    PreferredTimeZone = 'UTC'; PreferredSupportLanguage = 'en-us'
    RunStatePath = 'C:\test\state.json'; TicketTemplatePath = 'C:\test\template.json'
    TryAzCliToken = $true; UseDeviceCodeLogin = $false; RotateFingerprint = $true
    DryRun = $false; StopOnFirstFailure = $false; ResumeFromState = $false
    RetryFailedOnly = $true; SelectedRequestIds = @('id1', 'id2')
}

$v1Profile = [pscustomobject]@{
    profileVersion = 1; createdAt = '2026-01-01T00:00:00Z'; tokenSource = 'AzureCli'
    runSettings = @{
        DelaySeconds = 20; RequestsPerMinute = 10; MaxRetries = 6; BaseRetrySeconds = 25
        RotateFingerprint = $true; MaxRequests = 0; TryAzCliToken = $true
        UseDeviceCodeLogin = $false; StopOnFirstFailure = $false
    }
    execution = @{ DryRun = $false; RetryFailedRequests = $false; ResumeFromState = $false; CancelSignalPath = '' }
    proxy = @{ Url = ''; UseDefaultCredentials = $false }
    resume = @{ RunStatePath = 'C:\test\run-state.json' }
    defaults = @{ Region = 'eastus'; QuotaType = 'LowPriority' }
    ticket = @{
        NewLimit = 680; QuotaType = 'LowPriority'; Title = 'Quota request for Batch'
        ContactFirstName = 'Support'; ContactLastName = 'User'; ContactEmail = 'support@example.com'
        PreferredTimeZone = 'UTC'; PreferredSupportLanguage = 'en-us'
        TicketTemplatePath = 'C:\test\template.json'; ResultJsonPath = ''; ResultCsvPath = ''
    }
    ui = @{ SelectedRequestIds = @('id3') }
}

$hasMigrationFn = $null -ne (Get-Command Convert-ProfileToUnifiedSchema -ErrorAction SilentlyContinue)
Write-TestResult -Name 'Module exports Convert-ProfileToUnifiedSchema' -Passed $hasMigrationFn

if ($hasMigrationFn) {
    try {
        $result = Convert-ProfileToUnifiedSchema -Profile $flatProfile -TemplateDefaults $defaults
        Write-TestResult -Name 'Flat profile migration produces result' -Passed ($null -ne $result)
        Write-TestResult -Name 'Flat profile marked as migrated' -Passed ($result.Migrated -eq $true)
        Write-TestResult -Name 'Migrated profile has profileVersion=1' -Passed ($result.Profile.profileVersion -eq 1)
        Write-TestResult -Name 'Migrated profile has tokenSource' -Passed (-not [string]::IsNullOrWhiteSpace($result.Profile.tokenSource))
        Write-TestResult -Name 'Legacy RetryFailedOnly mapped to execution.RetryFailedRequests' -Passed ($result.Profile.execution.RetryFailedRequests -eq $true)
        Write-TestResult -Name 'Flat DelaySeconds preserved in runSettings' -Passed ($result.Profile.runSettings.DelaySeconds -eq 10)
        Write-TestResult -Name 'Flat NewLimit preserved in ticket section' -Passed ($result.Profile.ticket.NewLimit -eq 500)
        Write-TestResult -Name 'Flat SelectedRequestIds preserved in ui section' -Passed (@($result.Profile.ui.SelectedRequestIds).Count -eq 2)
    }
    catch {
        Write-TestResult -Name 'Flat profile migration' -Passed $false -Detail $_.Exception.Message
    }

    try {
        $result = Convert-ProfileToUnifiedSchema -Profile $v1Profile -TemplateDefaults $defaults
        Write-TestResult -Name 'V1 profile loads without migration flag' -Passed ($result.Migrated -eq $false)
        Write-TestResult -Name 'V1 profile preserves tokenSource' -Passed ($result.Profile.tokenSource -eq 'AzureCli')
        Write-TestResult -Name 'V1 profile preserves runSettings.DelaySeconds' -Passed ($result.Profile.runSettings.DelaySeconds -eq 20)
        Write-TestResult -Name 'V1 profile preserves resume.RunStatePath' -Passed ($result.Profile.resume.RunStatePath -eq 'C:\test\run-state.json')
    }
    catch {
        Write-TestResult -Name 'V1 profile loading' -Passed $false -Detail $_.Exception.Message
    }

    try {
        $mixedProfile = [pscustomobject]@{
            profileVersion = 1; createdAt = '2026-01-01T00:00:00Z'; tokenSource = 'AzureCli'
            runSettings = @{
                DelaySeconds = 15; RequestsPerMinute = 4; MaxRetries = 2; BaseRetrySeconds = 12
                RotateFingerprint = $false; MaxRequests = 0; TryAzCliToken = $true
                UseDeviceCodeLogin = $false; StopOnFirstFailure = $true
            }
            execution = @{ DryRun = $true; RetryFailedRequests = $true; ResumeFromState = $true; CancelSignalPath = '' }
            proxy = @{ Url = 'http://proxy:8080'; UseDefaultCredentials = $true }
            resume = @{ RunStatePath = 'C:\test\resume-state.json' }
            defaults = @{ Region = 'westus2'; QuotaType = 'Dedicated' }
            ticket = @{
                NewLimit = 1000; QuotaType = 'Dedicated'; Title = 'Dedicated quota request'
                ContactFirstName = 'Jane'; ContactLastName = 'Smith'; ContactEmail = 'jane@example.com'
                PreferredTimeZone = 'Pacific Standard Time'; PreferredSupportLanguage = 'en-us'
                TicketTemplatePath = 'C:\test\template.json'; ResultJsonPath = 'C:\test\result.json'; ResultCsvPath = 'C:\test\result.csv'
            }
            ui = @{ SelectedRequestIds = @('r1', 'r2', 'r3') }
        }
        $result = Convert-ProfileToUnifiedSchema -Profile $mixedProfile -TemplateDefaults $defaults
        Write-TestResult -Name 'Resume/retry profile preserves execution.ResumeFromState' -Passed ($result.Profile.execution.ResumeFromState -eq $true)
        Write-TestResult -Name 'Resume/retry profile preserves execution.RetryFailedRequests' -Passed ($result.Profile.execution.RetryFailedRequests -eq $true)
        Write-TestResult -Name 'Resume/retry profile preserves execution.DryRun' -Passed ($result.Profile.execution.DryRun -eq $true)
        Write-TestResult -Name 'Resume/retry profile preserves resume.RunStatePath' -Passed ($result.Profile.resume.RunStatePath -eq 'C:\test\resume-state.json')
        Write-TestResult -Name 'Resume/retry profile preserves proxy.Url' -Passed ($result.Profile.proxy.Url -eq 'http://proxy:8080')
        Write-TestResult -Name 'Resume/retry profile preserves defaults.Region' -Passed ($result.Profile.defaults.Region -eq 'westus2')
        Write-TestResult -Name 'Resume/retry profile preserves ticket.ResultJsonPath' -Passed ($result.Profile.ticket.ResultJsonPath -eq 'C:\test\result.json')
    }
    catch {
        Write-TestResult -Name 'Resume/retry profile migration' -Passed $false -Detail $_.Exception.Message
    }
}

# -------------------------------------------------------------------
Write-Host "`n=== Filter / Utility Parity ===" -ForegroundColor Cyan
# -------------------------------------------------------------------

$splitResult = @(Split-DiscoveryFilterList -Value 'sub1,sub2;sub3')
Write-TestResult -Name 'Split-DiscoveryFilterList handles comma+semicolon' -Passed ($splitResult.Count -eq 3)

$splitEmpty = @(Split-DiscoveryFilterList -Value '')
Write-TestResult -Name 'Split-DiscoveryFilterList returns empty for blank' -Passed ($splitEmpty.Count -eq 0)

$boolTrue = Convert-ToBoolValue -Value 'true' -Default $false
Write-TestResult -Name 'Convert-ToBoolValue parses "true"' -Passed ($boolTrue -eq $true)

$boolNull = Convert-ToBoolValue -Value $null -Default $true
Write-TestResult -Name 'Convert-ToBoolValue returns default for null' -Passed ($boolNull -eq $true)

$intVal = Convert-ToIntValue -Value '42' -Default 0
Write-TestResult -Name 'Convert-ToIntValue parses "42"' -Passed ($intVal -eq 42)

$intNull = Convert-ToIntValue -Value $null -Default 99
Write-TestResult -Name 'Convert-ToIntValue returns default for null' -Passed ($intNull -eq 99)

$arrVal = @(Convert-ToStringArray -Value @('a', '', 'b', $null, 'c'))
Write-TestResult -Name 'Convert-ToStringArray filters blanks and nulls' -Passed ($arrVal.Count -eq 3)

$firstDef = Get-FirstDefinedValue -Values @($null, $null, 'hello', 'world')
Write-TestResult -Name 'Get-FirstDefinedValue returns first non-null' -Passed ($firstDef -eq 'hello')

$objMember = Get-ObjectMemberValue -Object ([pscustomobject]@{ Foo = 'bar' }) -Name 'Foo'
Write-TestResult -Name 'Get-ObjectMemberValue reads PSObject property' -Passed ($objMember -eq 'bar')

$objMemberMissing = Get-ObjectMemberValue -Object ([pscustomobject]@{ Foo = 'bar' }) -Name 'Baz'
Write-TestResult -Name 'Get-ObjectMemberValue returns null for missing' -Passed ($null -eq $objMemberMissing)

# -------------------------------------------------------------------
Write-Host "`n=== Dry-Run Scenario: Template-Only ===" -ForegroundColor Cyan
# -------------------------------------------------------------------

$templateRequests = @()
if ($template.defaultRequests) {
    foreach ($dr in $template.defaultRequests) {
        $templateRequests += [pscustomobject]@{
            sub       = [string]$dr.sub
            account   = [string]$dr.account
            region    = [string]$dr.region
            limit     = if ($dr.PSObject.Properties['limit'] -and $null -ne $dr.limit) { [int]$dr.limit } else { [int]$defaults.NewLimit }
            quotaType = if ($dr.PSObject.Properties['quotaType'] -and -not [string]::IsNullOrWhiteSpace($dr.quotaType)) { [string]$dr.quotaType } else { [string]$defaults.QuotaType }
        }
    }
}
Write-TestResult -Name 'Template defaultRequests parsed' -Passed ($templateRequests.Count -gt 0) -Detail "Count: $($templateRequests.Count)"

if ($templateRequests.Count -gt 0) {
    $templatePreflight = Test-AzureSupportPreFlight `
        -Requests $templateRequests -Token 'template-dry-run-token' -TryAzCliToken $false `
        -DelaySeconds ([int]$defaults.DelaySeconds) -MaxRetries ([int]$defaults.MaxRetries) `
        -BaseRetrySeconds ([int]$defaults.BaseRetrySeconds) -RequestsPerMinute ([int]$defaults.RequestsPerMinute)
    Write-TestResult -Name 'Template-only preflight passes' -Passed $templatePreflight.IsValid -Detail "Errors: $($templatePreflight.Errors -join '; ')"

    foreach ($req in $templateRequests) {
        $hasRegion = -not [string]::IsNullOrWhiteSpace($req.region)
        Write-TestResult -Name "Template request '$($req.account)' has region" -Passed $hasRegion -Detail "Region: $($req.region)"
    }
}

# -------------------------------------------------------------------
Write-Host "`n=== Dry-Run Scenario: Autodiscovery Request Prep ===" -ForegroundColor Cyan
# -------------------------------------------------------------------

$mockDiscoveryRows = @(
    (New-DiscoveryGridRow -Id 'auto-1' -SubscriptionId 'sub-aaa' -SubscriptionName 'Sub A' -TenantId 'tenant-1' -AccountName 'batchacct1' -Region 'eastus' -DiscoveredRegions @('eastus', 'westus2')),
    (New-DiscoveryGridRow -Id 'auto-2' -SubscriptionId 'sub-aaa' -SubscriptionName 'Sub A' -TenantId 'tenant-1' -AccountName 'batchacct2' -Region 'westus2' -DiscoveredRegions @('westus2')),
    (New-DiscoveryGridRow -Id 'auto-3' -SubscriptionId 'sub-bbb' -SubscriptionName 'Sub B' -TenantId 'tenant-2' -AccountName 'batchacct3' -Region 'westeurope' -DiscoveredRegions @('westeurope', 'northeurope'))
)
Write-TestResult -Name 'Mock discovery rows created' -Passed ($mockDiscoveryRows.Count -eq 3)

$autoRequests = @()
foreach ($row in $mockDiscoveryRows) {
    $autoRequests += [pscustomobject]@{
        sub = [string]$row.SubscriptionId; account = [string]$row.AccountName
        region = [string]$row.Region; limit = [int]$defaults.NewLimit; quotaType = [string]$defaults.QuotaType
    }
}

$autoPreflight = Test-AzureSupportPreFlight `
    -Requests $autoRequests -Token 'auto-dry-run-token' -TryAzCliToken $false `
    -DelaySeconds ([int]$defaults.DelaySeconds) -MaxRetries ([int]$defaults.MaxRetries) `
    -BaseRetrySeconds ([int]$defaults.BaseRetrySeconds) -RequestsPerMinute ([int]$defaults.RequestsPerMinute)
Write-TestResult -Name 'Autodiscovery preflight passes' -Passed $autoPreflight.IsValid -Detail "Errors: $($autoPreflight.Errors -join '; ')"

foreach ($row in $mockDiscoveryRows) {
    $regionCheck = Test-DiscoveryRegionValue -Region $row.Region -DiscoveredRegions $row.DiscoveredRegions
    Write-TestResult -Name "Autodiscovery row '$($row.AccountName)' region valid" -Passed $regionCheck.IsValid
}

# -------------------------------------------------------------------
Write-Host "`n=== Dry-Run Scenario: Resume/Retry Profile Roundtrip ===" -ForegroundColor Cyan
# -------------------------------------------------------------------

$resumeProfilePath = Join-Path $testResultDir 'resume-profile-test.json'
$resumeProfile = [ordered]@{
    profileVersion = 1; createdAt = (Get-Date).ToString('o'); tokenSource = 'Token'
    runSettings = @{
        DelaySeconds = 5; RequestsPerMinute = 10; MaxRetries = 2; BaseRetrySeconds = 10
        RotateFingerprint = $true; MaxRequests = 0; TryAzCliToken = $false
        UseDeviceCodeLogin = $false; StopOnFirstFailure = $false
    }
    execution = @{ DryRun = $true; RetryFailedRequests = $true; ResumeFromState = $true; CancelSignalPath = '' }
    proxy = @{ Url = ''; UseDefaultCredentials = $false }
    resume = @{ RunStatePath = Join-Path $testResultDir 'test-run-state.json' }
    defaults = @{ Region = 'eastus'; QuotaType = 'LowPriority' }
    ticket = @{
        NewLimit = 500; QuotaType = 'LowPriority'; Title = 'Test resume ticket'
        ContactFirstName = 'Test'; ContactLastName = 'User'; ContactEmail = 'test@example.com'
        PreferredTimeZone = 'UTC'; PreferredSupportLanguage = 'en-us'
        TicketTemplatePath = $templatePath; ResultJsonPath = ''; ResultCsvPath = ''
    }
    ui = @{ SelectedRequestIds = @('auto-1', 'auto-2') }
}

try {
    Save-RunProfile -Path $resumeProfilePath -Profile $resumeProfile
    Write-TestResult -Name 'Resume profile saved to disk' -Passed (Test-Path -LiteralPath $resumeProfilePath)
}
catch {
    Write-TestResult -Name 'Resume profile saved to disk' -Passed $false -Detail $_.Exception.Message
}

try {
    $loadedProfile = Get-RunProfile -Path $resumeProfilePath
    Write-TestResult -Name 'Resume profile loaded from disk' -Passed ($null -ne $loadedProfile)
    $roundtripped = Convert-ProfileToUnifiedSchema -Profile $loadedProfile -TemplateDefaults $defaults
    Write-TestResult -Name 'Resume profile roundtrip not flagged as migrated' -Passed ($roundtripped.Migrated -eq $false)
    Write-TestResult -Name 'Resume profile roundtrip preserves RetryFailedRequests' -Passed ($roundtripped.Profile.execution.RetryFailedRequests -eq $true)
    Write-TestResult -Name 'Resume profile roundtrip preserves ResumeFromState' -Passed ($roundtripped.Profile.execution.ResumeFromState -eq $true)
    Write-TestResult -Name 'Resume profile roundtrip preserves DryRun' -Passed ($roundtripped.Profile.execution.DryRun -eq $true)
    Write-TestResult -Name 'Resume profile roundtrip preserves DelaySeconds' -Passed ($roundtripped.Profile.runSettings.DelaySeconds -eq 5)
}
catch {
    Write-TestResult -Name 'Resume profile roundtrip' -Passed $false -Detail $_.Exception.Message
}

if (Test-Path -LiteralPath $resumeProfilePath) {
    Remove-Item -LiteralPath $resumeProfilePath -Force -ErrorAction SilentlyContinue
}

# -------------------------------------------------------------------
Write-Host "`n=== Run State Timeline Resume ===" -ForegroundColor Cyan
# -------------------------------------------------------------------

class LegacyRunStateEntryForParity {
    [int]$index
    [string]$status
    [string]$subscription
    [string]$account
    [string]$region
    [int]$limit
    [string]$quotaType
    [object]$payload
    [string]$proxyUrl
    [string]$ticket
    [int]$attempts
    [int]$retryCount
    [double]$durationSeconds
    [string]$skipReason
    [string]$error
    [string]$startedAt
    [string]$completedAt
    [string]$timelineStartUtc
    [string]$timelineEndUtc
    [double]$interRequestDelaySeconds
    [double]$throttleWaitSeconds
    [string]$timeline
    [object[]]$attemptTimeline
}

$legacyEntry = [LegacyRunStateEntryForParity]::new()
$legacyEntry.index = 1
$legacyEntry.status = 'Failed'
$legacyEntry.subscription = 'sub-legacy'
$legacyEntry.account = 'acct-legacy'
$legacyEntry.region = 'eastus'
$legacyEntry.limit = 500
$legacyEntry.quotaType = 'LowPriority'
$legacyEntry.payload = $null
$legacyEntry.proxyUrl = ''
$legacyEntry.ticket = 'legacy-ticket'
$legacyEntry.attempts = 1
$legacyEntry.retryCount = 0
$legacyEntry.durationSeconds = 3.14
$legacyEntry.skipReason = ''
$legacyEntry.error = 'legacy failure'
$legacyEntry.startedAt = '2026-01-01T00:00:00Z'
$legacyEntry.completedAt = '2026-01-01T00:00:03Z'
$legacyEntry.timelineStartUtc = '2026-01-01T00:00:00Z'
$legacyEntry.timelineEndUtc = '2026-01-01T00:00:03Z'
$legacyEntry.interRequestDelaySeconds = 0
$legacyEntry.throttleWaitSeconds = 0
$legacyEntry.timeline = 'legacy-string-value'
$legacyEntry.attemptTimeline = @()

try {
    $normalizedEntry = & (Get-Module AzureSupport.TicketEngine) {
        param($Entry)
        Convert-ToWritableRunStateEntry -Entry $Entry
    } $legacyEntry
    Write-TestResult -Name 'Run state normalization returns PSCustomObject' -Passed ($normalizedEntry -is [pscustomobject])

    try {
        $normalizedEntry.timeline = [pscustomobject]@{
            attempts = @([pscustomobject]@{ attempt = 1; status = 'DryRun' })
        }
        Write-TestResult -Name 'Run state normalization allows structured timeline assignment' -Passed (@($normalizedEntry.timeline.attempts).Count -eq 1)
    }
    catch {
        Write-TestResult -Name 'Run state normalization allows structured timeline assignment' -Passed $false -Detail $_.Exception.Message
    }
}
catch {
    Write-TestResult -Name 'Run state normalization returns PSCustomObject' -Passed $false -Detail $_.Exception.Message
}

$resumeTimelineStatePath = Join-Path $testResultDir 'resume-timeline-run-state.json'
$resumeTimelineResultPath = Join-Path $testResultDir 'resume-timeline-results.json'
$resumeTimelineRequest = [pscustomobject]@{
    sub = 'sub-resume'
    account = 'acct-resume'
    region = 'eastus'
    limit = 500
    quotaType = 'LowPriority'
}

try {
    $resumeFingerprint = & (Get-Module AzureSupport.TicketEngine) {
        param($Requests)
        Get-RequestFingerprint -Requests $Requests
    } @($resumeTimelineRequest)

    $legacyResumeState = [ordered]@{
        runId = 'resume-timeline-test'
        status = 'Failed'
        requestedAction = 'Run'
        createdAt = '2026-01-01T00:00:00Z'
        updatedAt = '2026-01-01T00:00:03Z'
        lastError = 'legacy failure'
        stopOnFirstFailure = $false
        requestFingerprint = $resumeFingerprint
        runProfile = ''
        requestQueue = @(
            [ordered]@{
                index = 1
                status = 'Failed'
                subscription = 'sub-resume'
                account = 'acct-resume'
                region = 'eastus'
                limit = 500
                quotaType = 'LowPriority'
                payload = $null
                proxyUrl = ''
                ticket = 'legacy-ticket'
                attempts = 2
                retryCount = 1
                durationSeconds = 8.5
                skipReason = $null
                error = 'legacy failure'
                startedAt = '2026-01-01T00:00:00Z'
                completedAt = '2026-01-01T00:00:08Z'
                timelineStartUtc = '2026-01-01T00:00:00Z'
                timelineEndUtc = '2026-01-01T00:00:08Z'
                interRequestDelaySeconds = 0
                throttleWaitSeconds = 0
                timeline = [ordered]@{
                    startUtc = '2026-01-01T00:00:00Z'
                    endUtc = '2026-01-01T00:00:08Z'
                    durationSeconds = 8.5
                    interRequestDelaySeconds = 0
                    maxRetries = 1
                    retryCount = 1
                    throttleWaitSeconds = 0
                    attempts = @(
                        [ordered]@{
                            attempt = 1
                            status = 'Failed'
                            startedAtUtc = '2026-01-01T00:00:00Z'
                            endedAtUtc = '2026-01-01T00:00:08Z'
                            durationSeconds = 8.5
                            statusCode = 500
                            retryAfterSeconds = 0
                            sleepSeconds = 0
                            error = 'legacy failure'
                            reason = 'legacy request'
                        }
                    )
                }
                attemptTimeline = @(
                    [ordered]@{
                        attempt = 1
                        status = 'Failed'
                        startedAtUtc = '2026-01-01T00:00:00Z'
                        endedAtUtc = '2026-01-01T00:00:08Z'
                        durationSeconds = 8.5
                        statusCode = 500
                        retryAfterSeconds = 0
                        sleepSeconds = 0
                        error = 'legacy failure'
                        reason = 'legacy request'
                    }
                )
            }
        )
        totalRequests = 1
        completedRequests = 0
    }

    $legacyResumeState | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $resumeTimelineStatePath -Encoding UTF8

    $resumeResults = @(Invoke-AzureSupportBatchQuotaRun `
        -Requests @($resumeTimelineRequest) `
        -Token 'test-token' `
        -TryAzCliToken $false `
        -DryRun `
        -DelaySeconds 0 `
        -RequestsPerMinute 60 `
        -MaxRetries 2 `
        -BaseRetrySeconds 1 `
        -TicketTemplatePath $templatePath `
        -RunStatePath $resumeTimelineStatePath `
        -ResumeFromState $true `
        -RetryFailedRequests $true `
        -ResultJsonPath $resumeTimelineResultPath)

    Write-TestResult -Name 'Resume dry-run returns one result' -Passed ($resumeResults.Count -eq 1)
    Write-TestResult -Name 'Resume dry-run marks request DryRun' -Passed ($resumeResults.Count -eq 1 -and $resumeResults[0].status -eq 'DryRun')

    $savedResumeState = Get-Content -LiteralPath $resumeTimelineStatePath -Raw | ConvertFrom-Json
    $savedResumeEntry = @($savedResumeState.requestQueue)[0]
    Write-TestResult -Name 'Resume dry-run persists structured timeline in run state' -Passed ($null -ne $savedResumeEntry.timeline -and @($savedResumeEntry.timeline.attempts).Count -eq 1)
    Write-TestResult -Name 'Resume dry-run persists attempt timeline in run state' -Passed (@($savedResumeEntry.attemptTimeline).Count -eq 1)

    $savedResumeResults = @(Get-Content -LiteralPath $resumeTimelineResultPath -Raw | ConvertFrom-Json)
    Write-TestResult -Name 'Resume dry-run saves structured timeline in result JSON' -Passed ($savedResumeResults.Count -eq 1 -and $null -ne $savedResumeResults[0].timeline -and @($savedResumeResults[0].attemptTimeline).Count -eq 1)
}
catch {
    Write-TestResult -Name 'Run state timeline resume regression' -Passed $false -Detail $_.Exception.Message
}

foreach ($artifact in @($resumeTimelineStatePath, $resumeTimelineResultPath)) {
    if (Test-Path -LiteralPath $artifact) {
        Remove-Item -LiteralPath $artifact -Force -ErrorAction SilentlyContinue
    }
}

# -------------------------------------------------------------------
Write-Host "`n=== Summary ===" -ForegroundColor Cyan
# -------------------------------------------------------------------

Write-Host "  Total: $($pass + $fail)  Passed: $pass  Failed: $fail" -ForegroundColor $(if ($fail -eq 0) { 'Green' } else { 'Red' })

if ($testResultDir -and (Test-Path $testResultDir)) {
    $resultFile = Join-Path $testResultDir 'parity-results.json'
    $tests | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $resultFile -Encoding UTF8
    Write-Host "  Results saved to: $resultFile" -ForegroundColor Gray
}

if ($fail -gt 0) {
    exit 1
}
