$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$moduleRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$preferredEngineScriptPath = Join-Path $moduleRoot "Private\AzureSupport.TicketEngine.Core.ps1"
if (Test-Path -LiteralPath $preferredEngineScriptPath) {
    $script:EngineCoreScriptPath = (Resolve-Path -LiteralPath $preferredEngineScriptPath).Path
}
else {
    throw "Engine core script not found at '$preferredEngineScriptPath'."
}

function Get-ObjectMemberValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]$Object,
        [Parameter(Mandatory = $true)][string]$Name
    )

    if ($null -eq $Object) {
        return $null
    }

    if ($Object -is [System.Collections.IDictionary]) {
        if ($Object.Contains($Name)) {
            return $Object[$Name]
        }
        if ($Object.ContainsKey($Name)) {
            return $Object[$Name]
        }
        return $null
    }

    if ($null -eq $Object.PSObject) {
        return $null
    }

    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property) {
        return $null
    }

    return $property.Value
}

function Get-FirstDefinedValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object[]]$Values = @()
    )

    foreach ($value in $Values) {
        if ($null -ne $value) {
            return $value
        }
    }

    return $null
}

function Convert-ToBoolValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]$Value,
        [Parameter(Mandatory = $false)][bool]$Default = $false,
        [Parameter(Mandatory = $false)][string]$Name = "value",
        [Parameter(Mandatory = $false)][switch]$Strict
    )

    if ($null -eq $Value) {
        return $Default
    }

    if ($Value -is [bool] -or $Value -is [System.Management.Automation.SwitchParameter]) {
        return [bool]$Value
    }

    if ($Value -is [int]) {
        return $Value -ne 0
    }

    $normalized = [string]$Value
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return $Default
    }

    switch ($normalized.Trim().ToLowerInvariant()) {
        "true" { return $true }
        "false" { return $false }
        "1" { return $true }
        "0" { return $false }
        "yes" { return $true }
        "no" { return $false }
        "on" { return $true }
        "off" { return $false }
    }

    $shouldThrow = if ($PSBoundParameters.ContainsKey("Strict")) {
        [bool]$Strict
    }
    else {
        (-not [string]::IsNullOrWhiteSpace($Name) -and $Name -ne "value")
    }

    try {
        return [bool]$Value
    }
    catch {
        if ($shouldThrow) {
            throw "Invalid boolean value '$Value' for $Name."
        }
        return $Default
    }
}

function Convert-ToIntValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]$Value,
        [Parameter(Mandatory = $true)][int]$Default
    )

    if ($null -eq $Value) {
        return $Default
    }

    $parsed = 0
    if ([int]::TryParse([string]$Value, [ref]$parsed)) {
        return $parsed
    }

    return $Default
}

function Convert-ToStringArray {
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)]$Value)

    if ($null -eq $Value) {
        return @()
    }

    $values = New-Object System.Collections.Generic.List[string]
    foreach ($entry in @($Value)) {
        if ($null -eq $entry) {
            continue
        }

        if ($entry -is [System.Collections.IEnumerable] -and -not ($entry -is [string])) {
            foreach ($nestedEntry in @($entry)) {
                if ($null -eq $nestedEntry) {
                    continue
                }
                $nestedText = [string]$nestedEntry
                if (-not [string]::IsNullOrWhiteSpace($nestedText)) {
                    $values.Add($nestedText.Trim())
                }
            }
            continue
        }

        $text = [string]$entry
        if ([string]::IsNullOrWhiteSpace($text)) {
            continue
        }
        $values.Add($text.Trim())
    }

    return @($values)
}

function Get-DefaultStorageDirectory {
    $candidates = @($env:LOCALAPPDATA, $env:APPDATA, [System.IO.Path]::GetTempPath())
    $base = $candidates | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -First 1
    $storageDir = Join-Path $base "AzureSupportTickets"

    if (-not (Test-Path -LiteralPath $storageDir)) {
        try {
            New-Item -ItemType Directory -Path $storageDir -Force | Out-Null
        }
        catch {
            return [System.IO.Path]::GetTempPath()
        }
    }

    return $storageDir
}

function Resolve-DefaultedPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][string]$Path,
        [Parameter(Mandatory = $true)][string]$DefaultFileName
    )

    if (-not [string]::IsNullOrWhiteSpace($Path)) {
        return $Path
    }

    return Join-Path (Get-DefaultStorageDirectory) $DefaultFileName
}

function Get-RunProfile {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path)) {
        return $null
    }

    try {
        $raw = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
        if ([string]::IsNullOrWhiteSpace($raw)) {
            return $null
        }
        return ConvertFrom-Json -InputObject $raw -ErrorAction Stop
    }
    catch {
        Write-Warning "Unable to read run profile from '$Path': $($_.Exception.Message)"
        return $null
    }
}

function Save-RunProfile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)]$Profile
    )

    $parentPath = Split-Path -Parent $Path
    if (-not [string]::IsNullOrWhiteSpace($parentPath) -and -not (Test-Path -LiteralPath $parentPath)) {
        New-Item -ItemType Directory -Path $parentPath -Force | Out-Null
    }

    $Profile | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $Path -Encoding UTF8
}

function Get-TicketTemplatePath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][string]$TicketTemplatePath,
        [Parameter(Mandatory = $false)][string]$RootPath
    )

    if (-not [string]::IsNullOrWhiteSpace($TicketTemplatePath)) {
        return $TicketTemplatePath
    }

    $basePath = if (-not [string]::IsNullOrWhiteSpace($RootPath)) {
        $RootPath
    }
    elseif (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) {
        $PSScriptRoot
    }
    else {
        (Get-Location).Path
    }

    return Join-Path $basePath "config\default-ticket-template.json"
}

function Get-TicketTemplate {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Ticket template config not found at '$Path'."
    }

    try {
        $raw = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
        if ([string]::IsNullOrWhiteSpace($raw)) {
            throw "Ticket template is empty."
        }
        return ConvertFrom-Json -InputObject $raw -ErrorAction Stop
    }
    catch {
        throw "Unable to read ticket template from '$Path': $($_.Exception.Message)"
    }
}

function Merge-TemplateDefaults {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)]$Template)

    $templateDefaults = Get-ObjectMemberValue -Object $Template -Name "defaults"
    if ($null -eq $templateDefaults) {
        $templateDefaults = @{}
    }

    $contactDetails = Get-ObjectMemberValue -Object $Template -Name "contactDetails"

    $delaySeconds = Convert-ToIntValue -Value (Get-FirstDefinedValue @((Get-ObjectMemberValue -Object $templateDefaults -Name "delaySeconds"), 23)) -Default 23
    $requestsPerMinute = Convert-ToIntValue -Value (Get-FirstDefinedValue @((Get-ObjectMemberValue -Object $templateDefaults -Name "requestsPerMinute"), 2)) -Default 2
    $maxRetries = Convert-ToIntValue -Value (Get-FirstDefinedValue @((Get-ObjectMemberValue -Object $templateDefaults -Name "maxRetries"), 6)) -Default 6
    $baseRetrySeconds = Convert-ToIntValue -Value (Get-FirstDefinedValue @((Get-ObjectMemberValue -Object $templateDefaults -Name "baseRetrySeconds"), 25)) -Default 25
    $newLimit = Convert-ToIntValue -Value (Get-FirstDefinedValue @((Get-ObjectMemberValue -Object $templateDefaults -Name "newLimit"), (Get-ObjectMemberValue -Object $Template -Name "newLimit"), 680)) -Default 680

    $quotaType = [string](Get-FirstDefinedValue @((Get-ObjectMemberValue -Object $templateDefaults -Name "quotaType"), "LowPriority"))
    if ([string]::IsNullOrWhiteSpace($quotaType)) {
        $quotaType = "LowPriority"
    }

    return [ordered]@{
        DelaySeconds = $delaySeconds
        RequestsPerMinute = $requestsPerMinute
        MaxRetries = $maxRetries
        BaseRetrySeconds = $baseRetrySeconds
        NewLimit = $newLimit
        QuotaType = $quotaType
        Title = [string](Get-FirstDefinedValue @((Get-ObjectMemberValue -Object $Template -Name "title"), "Quota request for Batch"))
        ContactFirstName = [string](Get-FirstDefinedValue @((Get-ObjectMemberValue -Object $contactDetails -Name "firstName"), "Support"))
        ContactLastName = [string](Get-FirstDefinedValue @((Get-ObjectMemberValue -Object $contactDetails -Name "lastName"), "User"))
        ContactEmail = [string](Get-FirstDefinedValue @((Get-ObjectMemberValue -Object $contactDetails -Name "primaryEmailAddress"), "support@example.com"))
        PreferredTimeZone = [string](Get-FirstDefinedValue @((Get-ObjectMemberValue -Object $contactDetails -Name "preferredTimeZone"), "UTC"))
        PreferredSupportLanguage = [string](Get-FirstDefinedValue @((Get-ObjectMemberValue -Object $contactDetails -Name "preferredSupportLanguage"), "en-us"))
        PreferredContactMethod = [string](Get-FirstDefinedValue @((Get-ObjectMemberValue -Object $contactDetails -Name "preferredContactMethod"), "email"))
        Country = [string](Get-FirstDefinedValue @((Get-ObjectMemberValue -Object $contactDetails -Name "country"), "US"))
        ProblemClassificationId = [string](Get-FirstDefinedValue @((Get-ObjectMemberValue -Object $Template -Name "problemClassificationId"), ""))
        ServiceId = [string](Get-FirstDefinedValue @((Get-ObjectMemberValue -Object $Template -Name "serviceId"), ""))
        SupportPlanId = [string](Get-FirstDefinedValue @((Get-ObjectMemberValue -Object $Template -Name "supportPlanId"), ""))
    }
}

$functionNames = @(
    "Test-AzureSupportPreFlight",
    "Get-AzureSupportRequestKey",
    "New-AzureSupportRequestState",
    "New-AzureSupportRunState",
    "Get-AzureSupportRunState",
    "Save-AzureSupportRunState",
    "Write-Log",
    "Convert-ToRedactedLogMessage",
    "Ensure-RunArtifactFolder",
    "Read-SafeJsonFile",
    "Write-SafeJsonFile",
    "Get-ResolvedArtifactPath",
    "Get-RequestFingerprint",
    "New-RunProfileSnapshot",
    "New-RunStateSnapshot",
    "Read-RunStateSnapshot",
    "Update-RunStateCounters",
    "Write-RunStateSnapshot",
    "Get-PendingRequestStateItems",
    "ConvertTo-ProxyPoolEntry",
    "Resolve-ProxyPoolEntries",
    "Test-RunPreflight",
    "Resolve-TemplateTokens",
    "Get-EffectiveInterRequestDelaySeconds",
    "Get-ErrorResponseBody",
    "Get-ExceptionResponse",
    "Get-StatusCode",
    "Is-ThrottledResponse",
    "Get-RetryAfterSeconds",
    "Get-AzureRegionList",
    "Invoke-AzCommand",
    "Invoke-AzDeviceCodeLogin",
    "Get-AccessTokenFromAzCli",
    "Get-SubscriptionTenantMapFromAzCli",
    "Get-SubscriptionTenantIdFromAzCli",
    "Get-TenantIdFromUnauthorizedBody",
    "Resolve-AzCliTokenForSubscription",
    "Get-SubscriptionsFromAzCli",
    "Get-BatchRequestsFromAzCli",
    "Expand-AccountRegionRequests",
    "Resolve-RequestFieldValue",
    "Invoke-AzureSupportBatchQuotaRun",
    "Invoke-AzureSupportBatchQuotaRunQueued"
)

$tokens = $null
$parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($script:EngineCoreScriptPath, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors -and $parseErrors.Count -gt 0) {
    $errorText = [string]::Join("`n", ($parseErrors | ForEach-Object { $_.Message }))
    throw "Unable to parse engine core script '$($script:EngineCoreScriptPath)'. Errors: $errorText"
}

$functionMap = @{}
$requestedNames = New-Object 'System.Collections.Generic.HashSet[string]'
foreach ($name in $functionNames) {
    [void]$requestedNames.Add($name)
}

foreach ($node in $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)) {
    if ($requestedNames.Contains($node.Name) -and -not $functionMap.ContainsKey($node.Name)) {
        $functionMap[$node.Name] = $node.Extent.Text
    }
}

$missing = New-Object 'System.Collections.Generic.List[string]'
foreach ($name in $functionNames) {
    if (-not $functionMap.ContainsKey($name)) {
        $missing.Add($name)
    }
}
if ($missing.Count -gt 0) {
    throw "Engine core script '$($script:EngineCoreScriptPath)' is missing expected functions: $($missing -join ', ')"
}

foreach ($name in $functionNames) {
    Invoke-Expression $functionMap[$name]
}

$script:CachedAzureRegions = $null

function ConvertTo-DiscoveryCollection {
    param([Parameter(Mandatory = $true)]$InputObject)

    if ($null -eq $InputObject) {
        return @()
    }
    if ($InputObject.PSObject.Properties['value']) {
        return @($InputObject.value)
    }
    if ($InputObject -is [System.Object[]]) {
        if ($InputObject.Count -eq 1 -and $InputObject[0] -is [System.Object[]]) {
            return @($InputObject[0])
        }
        return @($InputObject)
    }

    return @($InputObject)
}

function Split-DiscoveryFilterList {
    param([Parameter(Mandatory = $false)][string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return @()
    }

    return @(
        ($Value -split '[,;\r\n]' |
            ForEach-Object { $_.Trim() } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Select-Object -Unique)
    )
}

function Get-AzureSupportDiscoveryRows {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][string]$SubscriptionFilter,
        [Parameter(Mandatory = $false)][string]$TenantFilter,
        [Parameter(Mandatory = $false)][string]$RegionFilter,
        [Parameter(Mandatory = $false)][string]$AccountFilter
    )

    $subscriptionFilterValues = @(Split-DiscoveryFilterList -Value $SubscriptionFilter)
    $tenantFilterValue = if ([string]::IsNullOrWhiteSpace($TenantFilter)) { '' } else { $TenantFilter.Trim() }
    $regionFilterValue = if ([string]::IsNullOrWhiteSpace($RegionFilter)) { '' } else { $RegionFilter.Trim() }
    $accountFilterValue = if ([string]::IsNullOrWhiteSpace($AccountFilter)) { '' } else { $AccountFilter.Trim() }

    $rawSubscriptions = Invoke-AzCommand -CommandArgs @("account", "list", "--all", "-o", "json")
    $parsedSubscriptions = ConvertFrom-Json -InputObject $rawSubscriptions -ErrorAction Stop
    $subscriptions = ConvertTo-DiscoveryCollection -InputObject $parsedSubscriptions

    $selectedSubscriptions = New-Object System.Collections.Generic.List[object]
    foreach ($subscription in $subscriptions) {
        if ($null -eq $subscription) {
            continue
        }

        $subscriptionId = [string]$subscription.id
        if ([string]::IsNullOrWhiteSpace($subscriptionId)) {
            continue
        }
        $subscriptionId = $subscriptionId.Trim()

        if ($subscriptionFilterValues.Count -gt 0 -and ($subscriptionFilterValues -notcontains $subscriptionId)) {
            continue
        }

        $tenantId = [string]$subscription.tenantId
        if (-not [string]::IsNullOrWhiteSpace($tenantFilterValue) -and -not ($tenantId -like "*$tenantFilterValue*")) {
            continue
        }

        $selectedSubscriptions.Add([pscustomobject]@{
            SubscriptionId = $subscriptionId
            SubscriptionName = if ([string]::IsNullOrWhiteSpace($subscription.name)) { $subscriptionId } else { [string]$subscription.name }
            TenantId = if ([string]::IsNullOrWhiteSpace($tenantId)) { '' } else { $tenantId.Trim() }
        })
    }

    if ($selectedSubscriptions.Count -eq 0) {
        return @()
    }

    $subscriptionLookup = @{}
    $subscriptionIds = New-Object System.Collections.Generic.List[string]
    foreach ($item in $selectedSubscriptions) {
        if (-not $subscriptionLookup.ContainsKey($item.SubscriptionId)) {
            $subscriptionLookup[$item.SubscriptionId] = $item
            $subscriptionIds.Add([string]$item.SubscriptionId)
        }
    }

    $discoveredRequests = @(Get-BatchRequestsFromAzCli -SubscriptionList @($subscriptionIds))
    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($request in $discoveredRequests) {
        $subscriptionId = [string](Resolve-RequestFieldValue -Request $request -FieldNames @('sub', 'subscription', 'subscriptionId', 'SubscriptionId', 'Subscription'))
        $accountName = [string](Resolve-RequestFieldValue -Request $request -FieldNames @('account', 'accountName', 'AccountName', 'name'))
        $region = [string](Resolve-RequestFieldValue -Request $request -FieldNames @('region', 'location'))

        if ([string]::IsNullOrWhiteSpace($subscriptionId) -or [string]::IsNullOrWhiteSpace($accountName)) {
            continue
        }

        if ([string]::IsNullOrWhiteSpace($region)) {
            $region = "eastus"
        }

        if (-not [string]::IsNullOrWhiteSpace($accountFilterValue) -and -not ($accountName -like "*$accountFilterValue*")) {
            continue
        }
        if (-not [string]::IsNullOrWhiteSpace($regionFilterValue) -and -not ($region -like "*$regionFilterValue*")) {
            continue
        }

        $regionList = New-Object System.Collections.Generic.List[string]
        $regionListValue = Resolve-RequestFieldValue -Request $request -FieldNames @('regionList')
        $regionCandidates = if ($null -ne $regionListValue) { @($regionListValue) } else { @($region) }
        foreach ($candidate in $regionCandidates) {
            if ($null -eq $candidate) {
                continue
            }

            $splitCandidates = @($candidate)
            if ($candidate -is [string]) {
                $splitCandidates = $candidate -split ","
            }
            elseif (-not ($candidate -is [System.Collections.IEnumerable])) {
                $splitCandidates = @("$candidate")
            }

            foreach ($entry in $splitCandidates) {
                if ($null -eq $entry) {
                    continue
                }

                $normalizedRegion = "$entry".Trim()
                if ([string]::IsNullOrWhiteSpace($normalizedRegion)) {
                    continue
                }
                if (-not ($regionList -contains $normalizedRegion)) {
                    $null = $regionList.Add($normalizedRegion)
                }
            }
        }

        if ($regionList.Count -eq 0) {
            $null = $regionList.Add($region.Trim())
        }

        $subscriptionMeta = $subscriptionLookup[$subscriptionId]
        $rows.Add([pscustomobject]@{
            Id = [guid]::NewGuid().ToString()
            Selected = $true
            SubscriptionId = $subscriptionId.Trim()
            SubscriptionName = if ($null -ne $subscriptionMeta) { [string]$subscriptionMeta.SubscriptionName } else { $subscriptionId.Trim() }
            TenantId = if ($null -ne $subscriptionMeta) { [string]$subscriptionMeta.TenantId } else { '' }
            AccountName = $accountName.Trim()
            Region = $region.Trim()
            DiscoveredRegions = @($regionList)
            Status = "Discovered"
        })
    }

    return @($rows | Sort-Object SubscriptionName, SubscriptionId, AccountName, Region)
}

function Test-DiscoveryRegionValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][string]$Region,
        [Parameter(Mandatory = $false)][string[]]$DiscoveredRegions = @(),
        [Parameter(Mandatory = $false)][string]$DefaultRegion = 'eastus'
    )

    $errors = New-Object System.Collections.Generic.List[string]
    $warnings = New-Object System.Collections.Generic.List[string]

    $normalized = if ([string]::IsNullOrWhiteSpace($Region)) { '' } else { $Region.Trim() }

    if ([string]::IsNullOrWhiteSpace($normalized)) {
        $normalized = $DefaultRegion
        $warnings.Add("Region was empty; defaulted to '$DefaultRegion'.")
    }

    $validDiscovered = @(Convert-ToStringArray -Value $DiscoveredRegions)
    if ($validDiscovered.Count -gt 0 -and ($validDiscovered -notcontains $normalized)) {
        $warnings.Add("Region '$normalized' is not among discovered regions: $($validDiscovered -join ', ').")
    }

    return [pscustomobject]@{
        Region   = $normalized
        IsValid  = ($errors.Count -eq 0)
        Errors   = @($errors)
        Warnings = @($warnings)
    }
}

function New-DiscoveryGridRow {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Id,
        [Parameter(Mandatory = $true)][string]$SubscriptionId,
        [Parameter(Mandatory = $false)][string]$SubscriptionName,
        [Parameter(Mandatory = $false)][AllowEmptyString()][string]$TenantId = '',
        [Parameter(Mandatory = $true)][string]$AccountName,
        [Parameter(Mandatory = $false)][AllowEmptyString()][string]$Region = '',
        [Parameter(Mandatory = $false)][string[]]$DiscoveredRegions = @()
    )

    $normalizedRegion = if ([string]::IsNullOrWhiteSpace($Region)) { 'eastus' } else { $Region.Trim() }
    $normalizedDiscoveredRegions = @(Convert-ToStringArray -Value $DiscoveredRegions)
    if ($null -eq $normalizedDiscoveredRegions -or $normalizedDiscoveredRegions.Count -eq 0) {
        $normalizedDiscoveredRegions = @($normalizedRegion)
    }

    [pscustomobject]@{
        Id                = $Id
        Selected          = $true
        SubscriptionId    = $SubscriptionId
        SubscriptionName  = if ([string]::IsNullOrWhiteSpace($SubscriptionName)) { $SubscriptionId } else { $SubscriptionName }
        TenantId          = $TenantId
        AccountName       = $AccountName
        Region            = $normalizedRegion
        DiscoveredRegions = @($normalizedDiscoveredRegions)
        Status            = 'Discovered'
    }
}

function Convert-ProfileToUnifiedSchema {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Profile,
        [Parameter(Mandatory = $false)]$TemplateDefaults = $null,
        [Parameter(Mandatory = $false)][string]$DefaultRunStatePath = '',
        [Parameter(Mandatory = $false)][string]$DefaultTicketTemplatePath = ''
    )

    if ($null -eq $TemplateDefaults) {
        $TemplateDefaults = [ordered]@{
            DelaySeconds             = 23
            RequestsPerMinute        = 2
            MaxRetries               = 6
            BaseRetrySeconds         = 25
            NewLimit                 = 680
            QuotaType                = 'LowPriority'
            Title                    = 'Quota request for Batch'
            ContactFirstName         = 'Support'
            ContactLastName          = 'User'
            ContactEmail             = 'support@example.com'
            PreferredTimeZone        = 'UTC'
            PreferredSupportLanguage = 'en-us'
        }
    }

    $runSettings = Get-ObjectMemberValue -Object $Profile -Name 'runSettings'
    $execution = Get-ObjectMemberValue -Object $Profile -Name 'execution'
    $proxy = Get-ObjectMemberValue -Object $Profile -Name 'proxy'
    $resume = Get-ObjectMemberValue -Object $Profile -Name 'resume'
    $defaults = Get-ObjectMemberValue -Object $Profile -Name 'defaults'
    $ticket = Get-ObjectMemberValue -Object $Profile -Name 'ticket'
    $ui = Get-ObjectMemberValue -Object $Profile -Name 'ui'

    $legacyRetryOnly = Get-ObjectMemberValue -Object $Profile -Name 'RetryFailedOnly'
    $existingVersion = Get-ObjectMemberValue -Object $Profile -Name 'profileVersion'
    $isUnified = ($null -ne $runSettings -or $null -ne $execution -or $null -ne $proxy -or $null -ne $resume -or $null -ne $defaults -or $null -ne $ticket -or $null -ne $ui)
    $needsMigration = (-not $isUnified) -or ($null -ne $legacyRetryOnly) -or ($null -eq $existingVersion) -or ($null -eq (Get-ObjectMemberValue -Object $Profile -Name 'tokenSource'))

    $effectiveRunSettings = if ($null -ne $runSettings) { $runSettings } else { $Profile }
    $effectiveExecution = if ($null -ne $execution) { $execution } else { $Profile }
    $effectiveProxy = if ($null -ne $proxy) { $proxy } else { $Profile }
    $effectiveResume = if ($null -ne $resume) { $resume } else { $Profile }
    $effectiveDefaults = if ($null -ne $defaults) { $defaults } else { $Profile }
    $effectiveTicket = if ($null -ne $ticket) { $ticket } else { $Profile }
    $effectiveUi = if ($null -ne $ui) { $ui } else { $Profile }

    $quotaTypeValue = [string](Get-FirstDefinedValue @(
            (Get-ObjectMemberValue -Object $effectiveDefaults -Name 'QuotaType'),
            (Get-ObjectMemberValue -Object $effectiveTicket -Name 'QuotaType'),
            (Get-ObjectMemberValue -Object $Profile -Name 'QuotaType'),
            $TemplateDefaults.QuotaType
        ))
    if ([string]::IsNullOrWhiteSpace($quotaTypeValue)) {
        $quotaTypeValue = [string]$TemplateDefaults.QuotaType
    }

    $tryAzCliTokenValue = Convert-ToBoolValue -Value (Get-FirstDefinedValue @((Get-ObjectMemberValue -Object $effectiveRunSettings -Name 'TryAzCliToken'), (Get-ObjectMemberValue -Object $Profile -Name 'TryAzCliToken'))) -Default $true
    $existingTokenSource = Get-ObjectMemberValue -Object $Profile -Name 'tokenSource'
    $tokenSourceValue = if (-not [string]::IsNullOrWhiteSpace([string]$existingTokenSource)) {
        [string]$existingTokenSource
    }
    elseif ($tryAzCliTokenValue) {
        'AzureCli'
    }
    else {
        'None'
    }

    $normalized = [ordered]@{
        profileVersion = 1
        createdAt      = [string](Get-FirstDefinedValue @((Get-ObjectMemberValue -Object $Profile -Name 'createdAt'), (Get-Date).ToString('o')))
        tokenSource    = $tokenSourceValue
        runSettings    = @{
            DelaySeconds       = Convert-ToIntValue -Value (Get-FirstDefinedValue @((Get-ObjectMemberValue -Object $effectiveRunSettings -Name 'DelaySeconds'), (Get-ObjectMemberValue -Object $Profile -Name 'DelaySeconds'))) -Default ([int]$TemplateDefaults.DelaySeconds)
            RequestsPerMinute  = [math]::Max(1, (Convert-ToIntValue -Value (Get-FirstDefinedValue @((Get-ObjectMemberValue -Object $effectiveRunSettings -Name 'RequestsPerMinute'), (Get-ObjectMemberValue -Object $Profile -Name 'RequestsPerMinute'))) -Default ([int]$TemplateDefaults.RequestsPerMinute)))
            MaxRetries         = [math]::Max(0, (Convert-ToIntValue -Value (Get-FirstDefinedValue @((Get-ObjectMemberValue -Object $effectiveRunSettings -Name 'MaxRetries'), (Get-ObjectMemberValue -Object $Profile -Name 'MaxRetries'))) -Default ([int]$TemplateDefaults.MaxRetries)))
            BaseRetrySeconds   = [math]::Max(1, (Convert-ToIntValue -Value (Get-FirstDefinedValue @((Get-ObjectMemberValue -Object $effectiveRunSettings -Name 'BaseRetrySeconds'), (Get-ObjectMemberValue -Object $Profile -Name 'BaseRetrySeconds'))) -Default ([int]$TemplateDefaults.BaseRetrySeconds)))
            RotateFingerprint  = Convert-ToBoolValue -Value (Get-FirstDefinedValue @((Get-ObjectMemberValue -Object $effectiveRunSettings -Name 'RotateFingerprint'), (Get-ObjectMemberValue -Object $Profile -Name 'RotateFingerprint'))) -Default $true
            MaxRequests        = [math]::Max(0, (Convert-ToIntValue -Value (Get-FirstDefinedValue @((Get-ObjectMemberValue -Object $effectiveRunSettings -Name 'MaxRequests'), 0)) -Default 0))
            TryAzCliToken      = $tryAzCliTokenValue
            UseDeviceCodeLogin = Convert-ToBoolValue -Value (Get-FirstDefinedValue @((Get-ObjectMemberValue -Object $effectiveRunSettings -Name 'UseDeviceCodeLogin'), (Get-ObjectMemberValue -Object $Profile -Name 'UseDeviceCodeLogin'))) -Default $false
            StopOnFirstFailure = Convert-ToBoolValue -Value (Get-FirstDefinedValue @((Get-ObjectMemberValue -Object $effectiveRunSettings -Name 'StopOnFirstFailure'), (Get-ObjectMemberValue -Object $Profile -Name 'StopOnFirstFailure'))) -Default $false
        }
        execution      = @{
            DryRun              = Convert-ToBoolValue -Value (Get-FirstDefinedValue @((Get-ObjectMemberValue -Object $effectiveExecution -Name 'DryRun'), (Get-ObjectMemberValue -Object $Profile -Name 'DryRun'))) -Default $false
            RetryFailedRequests = Convert-ToBoolValue -Value (Get-FirstDefinedValue @((Get-ObjectMemberValue -Object $effectiveExecution -Name 'RetryFailedRequests'), (Get-ObjectMemberValue -Object $Profile -Name 'RetryFailedRequests'), $legacyRetryOnly)) -Default $false
            ResumeFromState     = Convert-ToBoolValue -Value (Get-FirstDefinedValue @((Get-ObjectMemberValue -Object $effectiveExecution -Name 'ResumeFromState'), (Get-ObjectMemberValue -Object $Profile -Name 'ResumeFromState'))) -Default $false
            CancelSignalPath    = [string](Get-FirstDefinedValue @((Get-ObjectMemberValue -Object $effectiveExecution -Name 'CancelSignalPath'), ''))
        }
        proxy          = @{
            Url                   = [string](Get-FirstDefinedValue @((Get-ObjectMemberValue -Object $effectiveProxy -Name 'Url'), (Get-ObjectMemberValue -Object $effectiveProxy -Name 'ProxyUrl'), ''))
            UseDefaultCredentials = Convert-ToBoolValue -Value (Get-FirstDefinedValue @((Get-ObjectMemberValue -Object $effectiveProxy -Name 'UseDefaultCredentials'), (Get-ObjectMemberValue -Object $effectiveProxy -Name 'ProxyUseDefaultCredentials'))) -Default $false
        }
        resume         = @{
            RunStatePath = [string](Get-FirstDefinedValue @((Get-ObjectMemberValue -Object $effectiveResume -Name 'RunStatePath'), (Get-ObjectMemberValue -Object $Profile -Name 'RunStatePath'), $DefaultRunStatePath))
        }
        defaults       = @{
            Region    = [string](Get-FirstDefinedValue @((Get-ObjectMemberValue -Object $effectiveDefaults -Name 'Region'), 'eastus'))
            QuotaType = $quotaTypeValue
        }
        ticket         = @{
            NewLimit                 = [math]::Max(1, (Convert-ToIntValue -Value (Get-FirstDefinedValue @((Get-ObjectMemberValue -Object $effectiveTicket -Name 'NewLimit'), (Get-ObjectMemberValue -Object $Profile -Name 'NewLimit'))) -Default ([int]$TemplateDefaults.NewLimit)))
            QuotaType                = $quotaTypeValue
            Title                    = [string](Get-FirstDefinedValue @((Get-ObjectMemberValue -Object $effectiveTicket -Name 'Title'), (Get-ObjectMemberValue -Object $Profile -Name 'Title'), [string]$TemplateDefaults.Title))
            ContactFirstName         = [string](Get-FirstDefinedValue @((Get-ObjectMemberValue -Object $effectiveTicket -Name 'ContactFirstName'), (Get-ObjectMemberValue -Object $Profile -Name 'ContactFirstName'), [string]$TemplateDefaults.ContactFirstName))
            ContactLastName          = [string](Get-FirstDefinedValue @((Get-ObjectMemberValue -Object $effectiveTicket -Name 'ContactLastName'), (Get-ObjectMemberValue -Object $Profile -Name 'ContactLastName'), [string]$TemplateDefaults.ContactLastName))
            ContactEmail             = [string](Get-FirstDefinedValue @((Get-ObjectMemberValue -Object $effectiveTicket -Name 'ContactEmail'), (Get-ObjectMemberValue -Object $Profile -Name 'ContactEmail'), [string]$TemplateDefaults.ContactEmail))
            PreferredTimeZone        = [string](Get-FirstDefinedValue @((Get-ObjectMemberValue -Object $effectiveTicket -Name 'PreferredTimeZone'), (Get-ObjectMemberValue -Object $Profile -Name 'PreferredTimeZone'), [string]$TemplateDefaults.PreferredTimeZone))
            PreferredSupportLanguage = [string](Get-FirstDefinedValue @((Get-ObjectMemberValue -Object $effectiveTicket -Name 'PreferredSupportLanguage'), (Get-ObjectMemberValue -Object $Profile -Name 'PreferredSupportLanguage'), [string]$TemplateDefaults.PreferredSupportLanguage))
            TicketTemplatePath       = [string](Get-FirstDefinedValue @((Get-ObjectMemberValue -Object $effectiveTicket -Name 'TicketTemplatePath'), (Get-ObjectMemberValue -Object $Profile -Name 'TicketTemplatePath'), $DefaultTicketTemplatePath))
            ResultJsonPath           = [string](Get-FirstDefinedValue @((Get-ObjectMemberValue -Object $effectiveTicket -Name 'ResultJsonPath'), (Get-ObjectMemberValue -Object $Profile -Name 'ResultJsonPath'), ''))
            ResultCsvPath            = [string](Get-FirstDefinedValue @((Get-ObjectMemberValue -Object $effectiveTicket -Name 'ResultCsvPath'), (Get-ObjectMemberValue -Object $Profile -Name 'ResultCsvPath'), ''))
        }
        ui             = @{
            SelectedRequestIds = Convert-ToStringArray -Value (Get-FirstDefinedValue @((Get-ObjectMemberValue -Object $effectiveUi -Name 'SelectedRequestIds'), (Get-ObjectMemberValue -Object $Profile -Name 'SelectedRequestIds')))
        }
    }

    return [pscustomobject]@{
        Profile  = [pscustomobject]$normalized
        Migrated = $needsMigration
    }
}

function Invoke-AzureSupportTicketCli {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][hashtable]$BoundParameters,
        [Parameter(Mandatory = $false)][string]$ScriptRootOverride
    )

    if (-not (Test-Path -LiteralPath $script:EngineCoreScriptPath)) {
        throw "Engine core script not found at '$($script:EngineCoreScriptPath)'."
    }

    $invokeParams = @{}
    foreach ($entry in $BoundParameters.GetEnumerator()) {
        $invokeParams[$entry.Key] = $entry.Value
    }

    if (-not [string]::IsNullOrWhiteSpace($ScriptRootOverride) -and -not $invokeParams.ContainsKey("ScriptRootOverride")) {
        $invokeParams["ScriptRootOverride"] = $ScriptRootOverride
    }

    & $script:EngineCoreScriptPath @invokeParams
}

Export-ModuleMember -Function @(
    "Invoke-AzureSupportBatchQuotaRun",
    "Invoke-AzureSupportBatchQuotaRunQueued",
    "Invoke-AzureSupportTicketCli",
    "Get-AzureSupportDiscoveryRows",
    "Get-SubscriptionsFromAzCli",
    "Get-BatchRequestsFromAzCli",
    "Expand-AccountRegionRequests",
    "Resolve-RequestFieldValue",
    "Test-AzureSupportPreFlight",
    "New-RunProfileSnapshot",
    "Convert-ToBoolValue",
    "Convert-ToIntValue",
    "Convert-ToStringArray",
    "Get-ObjectMemberValue",
    "Get-FirstDefinedValue",
    "Get-DefaultStorageDirectory",
    "Resolve-DefaultedPath",
    "Get-RunProfile",
    "Save-RunProfile",
    "Get-TicketTemplatePath",
    "Get-TicketTemplate",
    "Merge-TemplateDefaults",
    "Split-DiscoveryFilterList",
    "ConvertTo-DiscoveryCollection",
    "New-DiscoveryGridRow",
    "Test-DiscoveryRegionValue",
    "Convert-ProfileToUnifiedSchema",
    "Get-AzureRegionList"
)
