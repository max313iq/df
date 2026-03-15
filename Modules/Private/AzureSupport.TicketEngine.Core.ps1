[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$Token = $env:AZURE_BEARER_TOKEN,

    [Parameter(Mandatory = $false)]
    [int]$DelaySeconds = 23,

    [Parameter(Mandatory = $false)]
    [string]$TicketTemplatePath,

    [Parameter(Mandatory = $false)]
    [string]$ContactFirstName,

    [Parameter(Mandatory = $false)]
    [string]$ContactLastName,

    [Parameter(Mandatory = $false)]
    [string]$PreferredContactMethod,

    [Parameter(Mandatory = $false)]
    [string]$PrimaryEmailAddress,

    [Parameter(Mandatory = $false)]
    [string]$PreferredTimeZone,

    [Parameter(Mandatory = $false)]
    [string]$Country,

    [Parameter(Mandatory = $false)]
    [string]$PreferredSupportLanguage,

    [Parameter(Mandatory = $false)]
    [string[]]$AdditionalEmailAddresses,

    [Parameter(Mandatory = $false)]
    [string]$AcceptLanguage,

    [Parameter(Mandatory = $false)]
    [string]$ProblemClassificationId,

    [Parameter(Mandatory = $false)]
    [string]$ServiceId,

    [Parameter(Mandatory = $false)]
    [string]$Severity,

    [Parameter(Mandatory = $false)]
    [string]$Title,

    [Parameter(Mandatory = $false)]
    [string]$DescriptionTemplate,

    [Parameter(Mandatory = $false)]
    [string]$AdvancedDiagnosticConsent,

    [Parameter(Mandatory = $false)]
    [Nullable[bool]]$Require24X7Response,

    [Parameter(Mandatory = $false)]
    [string]$SupportPlanId,

    [Parameter(Mandatory = $false)]
    [string]$QuotaChangeRequestVersion,

    [Parameter(Mandatory = $false)]
    [string]$QuotaChangeRequestSubType,

    [Parameter(Mandatory = $false)]
    [string]$QuotaRequestType,

    [Parameter(Mandatory = $false)]
    [Nullable[int]]$NewLimit,

    [Parameter(Mandatory = $false)]
    [switch]$DryRun,

    [Parameter(Mandatory = $false)]
    [int]$MaxRequests = 0,

    [Parameter(Mandatory = $false)]
    [string]$ProxyUrl,

    [Parameter(Mandatory = $false)]
    [switch]$ProxyUseDefaultCredentials,

    [Parameter(Mandatory = $false)]
    [pscredential]$ProxyCredential,

    [Parameter(Mandatory = $false)]
    [string[]]$ProxyPool,

    [Parameter(Mandatory = $false)]
    [int]$MaxRetries = 6,

    [Parameter(Mandatory = $false)]
    [int]$BaseRetrySeconds = 25,

    [Parameter(Mandatory = $false)]
    [switch]$RotateFingerprint = $true,

    [Parameter(Mandatory = $false)]
    [switch]$AutoDiscoverRequests,

    [Parameter(Mandatory = $false)]
    [string[]]$SubscriptionIds,

    [Parameter(Mandatory = $false)]
    [object]$TryAzCliToken = $true,

    [Parameter(Mandatory = $false)]
    [object]$UseDeviceCodeLogin = $false,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 120)]
    [int]$RequestsPerMinute = 2,

    [Parameter(Mandatory = $false)]
    [string]$ResultJsonPath,

    [Parameter(Mandatory = $false)]
    [string]$ResultCsvPath,

    [Parameter(Mandatory = $false)]
    [string]$RunProfilePath,

    [Parameter(Mandatory = $false)]
    [object]$LoadRunProfile = $true,

    [Parameter(Mandatory = $false)]
    [object]$SaveRunProfile = $false,

    [Parameter(Mandatory = $false)]
    [string]$RunStatePath,

    [Parameter(Mandatory = $false)]
    [object]$ResumeFromState = $false,

    [Parameter(Mandatory = $false)]
    [object]$RetryFailedRequests = $false,

    [Parameter(Mandatory = $false)]
    [string]$CancelSignalPath,

    [Parameter(Mandatory = $false)]
    [string]$ScriptRootOverride,

    [Parameter(Mandatory = $false)]
    [switch]$StopOnFirstFailure
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Convert-ToBoolValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]$Value,
        [Parameter(Mandatory = $false)][bool]$Default = $false,
        [Parameter(Mandatory = $false)][string]$Name = "value"
    )

    return AzureSupport.TicketEngine\Convert-ToBoolValue -Value $Value -Default $Default -Name $Name
}

function Remove-EmptyParameters {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Parameters
    )

    $clean = @{}
    foreach ($entry in $Parameters.GetEnumerator()) {
        $value = $entry.Value
        if ($entry.Key -eq "Token") {
            $clean[$entry.Key] = [string]$value
            continue
        }
        if ($null -eq $value) {
            continue
        }
        if ($value -is [string] -and [string]::IsNullOrWhiteSpace($value)) {
            continue
        }
        $clean[$entry.Key] = $value
    }
    return $clean
}

function Get-SafeProfileProperty {
    param(
        [Parameter(Mandatory = $false)]$Object,
        [Parameter(Mandatory = $true)][string]$Name
    )

    return AzureSupport.TicketEngine\Get-ObjectMemberValue -Object $Object -Name $Name
}

$TryAzCliToken = Convert-ToBoolValue -Value $TryAzCliToken -Default $true -Name "TryAzCliToken"
$UseDeviceCodeLogin = Convert-ToBoolValue -Value $UseDeviceCodeLogin -Default $false -Name "UseDeviceCodeLogin"
$LoadRunProfile = Convert-ToBoolValue -Value $LoadRunProfile -Default $true -Name "LoadRunProfile"
$SaveRunProfile = Convert-ToBoolValue -Value $SaveRunProfile -Default $false -Name "SaveRunProfile"

$scriptRootForDefaults = if (-not [string]::IsNullOrWhiteSpace($ScriptRootOverride)) { $ScriptRootOverride } elseif (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) { $PSScriptRoot } else { (Get-Location).Path }
$RunProfilePath = if ([string]::IsNullOrWhiteSpace($RunProfilePath)) { Join-Path $scriptRootForDefaults "azure-ticket-run-profile.json" } else { $RunProfilePath }
$RunStatePath = if ([string]::IsNullOrWhiteSpace($RunStatePath)) { Join-Path $scriptRootForDefaults "azure-support-ticket-run-state.json" } else { $RunStatePath }

if ($LoadRunProfile) {
    $savedProfile = $null
    if (Test-Path -PathType Leaf $RunProfilePath) {
        try {
            $rawProfile = Get-Content -Raw -Path $RunProfilePath -ErrorAction Stop
            if (-not [string]::IsNullOrWhiteSpace($rawProfile)) {
                $savedProfile = ConvertFrom-Json -InputObject $rawProfile -ErrorAction Stop
            }
        }
        catch {
            Write-Warning "Unable to load run profile from '$RunProfilePath': $($_.Exception.Message)"
            $savedProfile = $null
        }
    }
    if ($null -ne $savedProfile) {
        $runSettings = Get-SafeProfileProperty -Object $savedProfile -Name 'runSettings'
        $executionSettings = Get-SafeProfileProperty -Object $savedProfile -Name 'execution'
        $proxySettings = Get-SafeProfileProperty -Object $savedProfile -Name 'proxy'
        $resumeSettings = Get-SafeProfileProperty -Object $savedProfile -Name 'resume'

        $value = Get-SafeProfileProperty -Object $runSettings -Name 'DelaySeconds'
        if ($null -eq $value) { $value = Get-SafeProfileProperty -Object $savedProfile -Name 'DelaySeconds' }
        if (-not $PSBoundParameters.ContainsKey('DelaySeconds') -and $null -ne $value) { $DelaySeconds = [int]$value }

        $value = Get-SafeProfileProperty -Object $runSettings -Name 'MaxRequests'
        if ($null -eq $value) { $value = Get-SafeProfileProperty -Object $savedProfile -Name 'MaxRequests' }
        if (-not $PSBoundParameters.ContainsKey('MaxRequests') -and $null -ne $value) { $MaxRequests = [int]$value }

        $value = Get-SafeProfileProperty -Object $runSettings -Name 'MaxRetries'
        if ($null -eq $value) { $value = Get-SafeProfileProperty -Object $savedProfile -Name 'MaxRetries' }
        if (-not $PSBoundParameters.ContainsKey('MaxRetries') -and $null -ne $value) { $MaxRetries = [int]$value }

        $value = Get-SafeProfileProperty -Object $runSettings -Name 'BaseRetrySeconds'
        if ($null -eq $value) { $value = Get-SafeProfileProperty -Object $savedProfile -Name 'BaseRetrySeconds' }
        if (-not $PSBoundParameters.ContainsKey('BaseRetrySeconds') -and $null -ne $value) { $BaseRetrySeconds = [int]$value }

        $value = Get-SafeProfileProperty -Object $runSettings -Name 'RequestsPerMinute'
        if ($null -eq $value) { $value = Get-SafeProfileProperty -Object $savedProfile -Name 'RequestsPerMinute' }
        if (-not $PSBoundParameters.ContainsKey('RequestsPerMinute') -and $null -ne $value) { $RequestsPerMinute = [int]$value }

        $value = Get-SafeProfileProperty -Object $runSettings -Name 'RotateFingerprint'
        if ($null -eq $value) { $value = Get-SafeProfileProperty -Object $savedProfile -Name 'RotateFingerprint' }
        if (-not $PSBoundParameters.ContainsKey('RotateFingerprint') -and $null -ne $value) { $RotateFingerprint = $value }

        $value = Get-SafeProfileProperty -Object $runSettings -Name 'TryAzCliToken'
        if ($null -eq $value) { $value = Get-SafeProfileProperty -Object $savedProfile -Name 'TryAzCliToken' }
        if (-not $PSBoundParameters.ContainsKey('TryAzCliToken') -and $null -ne $value) { $TryAzCliToken = $value }

        $value = Get-SafeProfileProperty -Object $runSettings -Name 'UseDeviceCodeLogin'
        if ($null -eq $value) { $value = Get-SafeProfileProperty -Object $savedProfile -Name 'UseDeviceCodeLogin' }
        if (-not $PSBoundParameters.ContainsKey('UseDeviceCodeLogin') -and $null -ne $value) { $UseDeviceCodeLogin = $value }

        $value = Get-SafeProfileProperty -Object $runSettings -Name 'StopOnFirstFailure'
        if ($null -eq $value) { $value = Get-SafeProfileProperty -Object $savedProfile -Name 'StopOnFirstFailure' }
        if (-not $PSBoundParameters.ContainsKey('StopOnFirstFailure') -and $null -ne $value) { $StopOnFirstFailure = $value }

        $value = Get-SafeProfileProperty -Object $executionSettings -Name 'DryRun'
        if ($null -eq $value) { $value = Get-SafeProfileProperty -Object $savedProfile -Name 'DryRun' }
        if (-not $PSBoundParameters.ContainsKey('DryRun') -and $null -ne $value) { $DryRun = $value }

        $value = Get-SafeProfileProperty -Object $runSettings -Name 'AutoDiscoverRequests'
        if ($null -eq $value) { $value = Get-SafeProfileProperty -Object $savedProfile -Name 'AutoDiscoverRequests' }
        if (-not $PSBoundParameters.ContainsKey('AutoDiscoverRequests') -and $null -ne $value) { $AutoDiscoverRequests = $value }

        $value = Get-SafeProfileProperty -Object $proxySettings -Name 'Url'
        if ($null -eq $value) { $value = Get-SafeProfileProperty -Object $runSettings -Name 'ProxyUrl' }
        if ($null -eq $value) { $value = Get-SafeProfileProperty -Object $savedProfile -Name 'ProxyUrl' }
        if (-not $PSBoundParameters.ContainsKey('ProxyUrl') -and $null -ne $value) { $ProxyUrl = [string]$value }

        $value = Get-SafeProfileProperty -Object $proxySettings -Name 'UseDefaultCredentials'
        if ($null -eq $value) { $value = Get-SafeProfileProperty -Object $runSettings -Name 'ProxyUseDefaultCredentials' }
        if ($null -eq $value) { $value = Get-SafeProfileProperty -Object $savedProfile -Name 'ProxyUseDefaultCredentials' }
        if (-not $PSBoundParameters.ContainsKey('ProxyUseDefaultCredentials') -and $null -ne $value) { $ProxyUseDefaultCredentials = $value }

        $value = Get-SafeProfileProperty -Object $proxySettings -Name 'Pool'
        if ($null -eq $value) { $value = Get-SafeProfileProperty -Object $runSettings -Name 'ProxyPool' }
        if ($null -eq $value) { $value = Get-SafeProfileProperty -Object $savedProfile -Name 'ProxyPool' }
        if (-not $PSBoundParameters.ContainsKey('ProxyPool') -and $null -ne $value) {
            if ($value -is [array]) {
                $ProxyPool = @($value)
            }
            elseif (-not [string]::IsNullOrWhiteSpace([string]$value)) {
                $ProxyPool = @([string]$value)
            }
        }

        $value = Get-SafeProfileProperty -Object $runSettings -Name 'ResultJsonPath'
        if ($null -eq $value) { $value = Get-SafeProfileProperty -Object $savedProfile -Name 'ResultJsonPath' }
        if (-not $PSBoundParameters.ContainsKey('ResultJsonPath') -and $null -ne $value) { $ResultJsonPath = [string]$value }

        $value = Get-SafeProfileProperty -Object $resumeSettings -Name 'RunStatePath'
        if ($null -eq $value) { $value = Get-SafeProfileProperty -Object $runSettings -Name 'RunStatePath' }
        if ($null -eq $value) { $value = Get-SafeProfileProperty -Object $savedProfile -Name 'RunStatePath' }
        if (-not $PSBoundParameters.ContainsKey('RunStatePath') -and $null -ne $value) { $RunStatePath = [string]$value }

        $value = Get-SafeProfileProperty -Object $executionSettings -Name 'ResumeFromState'
        if ($null -eq $value) { $value = Get-SafeProfileProperty -Object $savedProfile -Name 'ResumeFromState' }
        if (-not $PSBoundParameters.ContainsKey('ResumeFromState') -and $null -ne $value) { $ResumeFromState = $value }

        $value = Get-SafeProfileProperty -Object $executionSettings -Name 'RetryFailedRequests'
        if ($null -eq $value) { $value = Get-SafeProfileProperty -Object $savedProfile -Name 'RetryFailedRequests' }
        if (-not $PSBoundParameters.ContainsKey('RetryFailedRequests') -and $null -ne $value) { $RetryFailedRequests = $value }

        $value = Get-SafeProfileProperty -Object $executionSettings -Name 'CancelSignalPath'
        if ($null -eq $value) { $value = Get-SafeProfileProperty -Object $savedProfile -Name 'CancelSignalPath' }
        if (-not $PSBoundParameters.ContainsKey('CancelSignalPath') -and $null -ne $value) { $CancelSignalPath = [string]$value }

        $value = Get-SafeProfileProperty -Object $savedProfile -Name 'SubscriptionIds'
        if ($null -ne $value -and -not $PSBoundParameters.ContainsKey('SubscriptionIds')) {
            if ($value -is [array]) {
                $SubscriptionIds = @($value)
            }
            elseif (-not [string]::IsNullOrWhiteSpace([string]$value)) {
                $SubscriptionIds = @([string]$value)
            }
        }
    }
}

$RotateFingerprint = Convert-ToBoolValue -Value $RotateFingerprint -Default $true -Name "RotateFingerprint"
$TryAzCliToken = Convert-ToBoolValue -Value $TryAzCliToken -Default $true -Name "TryAzCliToken"
$UseDeviceCodeLogin = Convert-ToBoolValue -Value $UseDeviceCodeLogin -Default $false -Name "UseDeviceCodeLogin"
$LoadRunProfile = Convert-ToBoolValue -Value $LoadRunProfile -Default $true -Name "LoadRunProfile"
$SaveRunProfile = Convert-ToBoolValue -Value $SaveRunProfile -Default $false -Name "SaveRunProfile"
$AutoDiscoverRequests = Convert-ToBoolValue -Value $AutoDiscoverRequests -Default $false -Name "AutoDiscoverRequests"
$ProxyUseDefaultCredentials = Convert-ToBoolValue -Value $ProxyUseDefaultCredentials -Default $false -Name "ProxyUseDefaultCredentials"
$DryRun = Convert-ToBoolValue -Value $DryRun -Default $false -Name "DryRun"
$StopOnFirstFailure = Convert-ToBoolValue -Value $StopOnFirstFailure -Default $false -Name "StopOnFirstFailure"
$ResumeFromState = Convert-ToBoolValue -Value $ResumeFromState -Default $false -Name "ResumeFromState"
$RetryFailedRequests = Convert-ToBoolValue -Value $RetryFailedRequests -Default $false -Name "RetryFailedRequests"
if ($null -eq $ProxyPool) {
    $ProxyPool = @()
}
else {
    $ProxyPool = @(
        @($ProxyPool) |
            ForEach-Object { if ($null -eq $_) { '' } else { [string]$_ } } |
            ForEach-Object { $_.Trim() } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    )
}

function Get-DefaultStorageDirectory {
    return AzureSupport.TicketEngine\Get-DefaultStorageDirectory
}

function Resolve-DefaultedPath {
    param(
        [Parameter(Mandatory = $false)][string]$Path,
        [Parameter(Mandatory = $true)][string]$DefaultFileName
    )

    return AzureSupport.TicketEngine\Resolve-DefaultedPath -Path $Path -DefaultFileName $DefaultFileName
}

function Get-RunProfile {
    param([Parameter(Mandatory = $true)][string]$Path)
    return AzureSupport.TicketEngine\Get-RunProfile -Path $Path
}

function Save-RunProfile {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)]$Profile
    )
    AzureSupport.TicketEngine\Save-RunProfile -Path $Path -Profile $Profile
}

function Test-AzureSupportPreFlight {
    param(
        [Parameter(Mandatory = $true)][array]$Requests,
        [Parameter(Mandatory = $true)][string]$Token,
        [Parameter(Mandatory = $false)][bool]$TryAzCliToken = $true,
        [Parameter(Mandatory = $false)][bool]$UseDeviceCodeLogin = $false,
        [Parameter(Mandatory = $false)][int]$DelaySeconds = 23,
        [Parameter(Mandatory = $false)][int]$MaxRequests = 0,
        [Parameter(Mandatory = $false)][int]$MaxRetries = 6,
        [Parameter(Mandatory = $false)][int]$BaseRetrySeconds = 25,
        [Parameter(Mandatory = $false)][bool]$RotateFingerprint = $true,
        [Parameter(Mandatory = $false)][int]$RequestsPerMinute = 2,
        [Parameter(Mandatory = $false)][string]$ProxyUrl,
        [Parameter(Mandatory = $false)][bool]$ProxyUseDefaultCredentials = $false,
        [Parameter(Mandatory = $false)][pscredential]$ProxyCredential,
        [Parameter(Mandatory = $false)][string]$RunStatePath,
        [Parameter(Mandatory = $false)][bool]$StopOnFirstFailure = $false
    )

    $errors = New-Object System.Collections.Generic.List[string]
    $warnings = New-Object System.Collections.Generic.List[string]

    if ($MaxRequests -lt 0) { $errors.Add("MaxRequests cannot be negative.") }
    if ($MaxRetries -lt 0) { $errors.Add("MaxRetries cannot be negative.") }
    if ($BaseRetrySeconds -lt 1) { $errors.Add("BaseRetrySeconds must be >= 1.") }
    if ($DelaySeconds -lt 0) { $errors.Add("DelaySeconds cannot be negative.") }
    if (-not $UseDeviceCodeLogin -and $DelaySeconds -gt 120) {
        $warnings.Add("DelaySeconds above 120 seconds may lengthen total runtime.")
    }

    if (-not $Requests -or $Requests.Count -eq 0) {
        $errors.Add("No quota requests were provided.")
    }

    $requestKeys = New-Object System.Collections.Generic.HashSet[string]
    $validQuotaTypes = @("LowPriority", "Dedicated", "Standard", "Spot")
    $index = 0
    foreach ($request in $Requests) {
        $index++
        $subscription = Resolve-RequestFieldValue -Request $request -FieldNames @('sub', 'subscription', 'subscriptionId', 'SubscriptionId', 'Subscription')
        $account = Resolve-RequestFieldValue -Request $request -FieldNames @('account', 'accountName', 'AccountName', 'name')
        $region = Resolve-RequestFieldValue -Request $request -FieldNames @('region', 'location')
        $limitRaw = Resolve-RequestFieldValue -Request $request -FieldNames @('newLimit', 'NewLimit', 'limit')
        $quotaType = Resolve-RequestFieldValue -Request $request -FieldNames @('quotaType', 'type', 'Type')

        if ([string]::IsNullOrWhiteSpace($subscription)) {
            $errors.Add("Request #$index is missing required subscription.")
            continue
        }
        if ([string]::IsNullOrWhiteSpace($account)) {
            $errors.Add("Request #$index is missing required account.")
            continue
        }
        if ([string]::IsNullOrWhiteSpace($region)) {
            $region = "eastus"
        }

        if ($null -ne $limitRaw) {
            $limitParsed = 0
            if (-not [int]::TryParse([string]$limitRaw, [ref]$limitParsed)) {
                $errors.Add("Request #$index has non-numeric newLimit '$limitRaw'.")
            }
            elseif ($limitParsed -le 0) {
                $errors.Add("Request #$index has non-positive newLimit '$limitParsed'.")
            }
        }

        if (-not [string]::IsNullOrWhiteSpace($quotaType)) {
            if ($validQuotaTypes -notcontains $quotaType) {
                $warnings.Add("Request #$index has quotaType '$quotaType'. Validated values are: $($validQuotaTypes -join ', ').")
            }
        }

        $requestKey = "${subscription}|${account}|${region}"
        if ($requestKeys.Contains($requestKey)) {
            $errors.Add("Duplicate request found for '$requestKey'. Use unique subscription/account/region combinations.")
        }
        else {
            [void]$requestKeys.Add($requestKey)
        }
    }

    if ([string]::IsNullOrWhiteSpace($Token)) {
        if (-not $TryAzCliToken) {
            $errors.Add("Token was not provided and -TryAzCliToken is false.")
        }
        else {
            if (-not (Get-Command az -ErrorAction SilentlyContinue)) {
                $errors.Add("Token was not provided and Azure CLI was not found, but -TryAzCliToken is enabled.")
            }
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($ProxyUrl)) {
        if ($ProxyUrl -notmatch '^(http|https)://') {
            $warnings.Add("ProxyUrl '$ProxyUrl' does not start with http:// or https://.")
        }
        if ($ProxyUseDefaultCredentials -and $ProxyCredential) {
            $warnings.Add("Both ProxyUseDefaultCredentials and ProxyCredential are set. Default credentials will be used.")
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($RunStatePath)) {
        $runStateDir = Split-Path -Parent $RunStatePath
        if (-not [string]::IsNullOrWhiteSpace($runStateDir) -and -not (Test-Path $runStateDir)) {
            try {
                New-Item -ItemType Directory -Path $runStateDir -Force | Out-Null
            }
            catch {
                $warnings.Add("RunStatePath directory '$runStateDir' is not writable.")
            }
        }
    }

    return [pscustomobject]@{
        IsValid = $errors.Count -eq 0
        Errors = $errors
        Warnings = $warnings
        RotatingFingerprintEnabled = $RotateFingerprint
        StopOnFirstFailure = $StopOnFirstFailure
    }
}

function Get-AzureSupportRequestKey {
    param([Parameter(Mandatory = $true)]$Request)

    $sub = if ([string]::IsNullOrWhiteSpace($Request.sub)) { "" } else { $Request.sub.Trim() }
    $account = if ([string]::IsNullOrWhiteSpace($Request.account)) { "" } else { $Request.account.Trim() }
    $region = if ([string]::IsNullOrWhiteSpace($Request.region)) { "" } else { $Request.region.Trim() }

    return "${sub}|${account}|${region}"
}

function New-AzureSupportRequestState {
    param([Parameter(Mandatory = $true)]$Request)

    return [ordered]@{
        requestKey = Get-AzureSupportRequestKey -Request $Request
        index = $Request.Index
        account = $Request.account
        subscription = $Request.sub
        region = $Request.region
        ticket = $null
        status = "Pending"
        attempts = 0
        durationSeconds = 0.0
        error = $null
        startedAt = $null
        completedAt = $null
    }
}

function New-AzureSupportRunState {
    param([Parameter(Mandatory = $true)][array]$ValidatedRequests)

    $entries = New-Object System.Collections.Generic.List[object]
    foreach ($request in $ValidatedRequests) {
        $entries.Add((New-AzureSupportRequestState -Request $request))
    }

    return [ordered]@{
        runId = [guid]::NewGuid().ToString()
        version = 1
        status = "Running"
        startedAt = (Get-Date).ToString("o")
        lastUpdatedAt = (Get-Date).ToString("o")
        requests = @($entries)
    }
}

function Get-AzureSupportRunState {
    param([Parameter(Mandatory = $true)][string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path $Path)) {
        return $null
    }

    try {
        $raw = Get-Content -Path $Path -Raw -ErrorAction Stop
        if ([string]::IsNullOrWhiteSpace($raw)) {
            return $null
        }
        return ConvertFrom-Json -InputObject $raw -ErrorAction Stop
    }
    catch {
        Write-Warning "Unable to read run state from '$Path': $($_.Exception.Message)"
        return $null
    }
}

function Save-AzureSupportRunState {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)]$State
    )

    if ([string]::IsNullOrWhiteSpace($Path)) { return }

    $parentPath = Split-Path -Parent $Path
    if (-not [string]::IsNullOrWhiteSpace($parentPath) -and -not (Test-Path $parentPath)) {
        New-Item -ItemType Directory -Path $parentPath -Force | Out-Null
    }

    $State.lastUpdatedAt = (Get-Date).ToString("o")
    $State | ConvertTo-Json -Depth 12 | Set-Content -Path $Path -Encoding UTF8
}

function Write-Log {
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [Parameter(Mandatory = $false)][ValidateSet('INFO', 'WARN', 'ERROR')][string]$Level = 'INFO'
    )

    $stamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    $safeMessage = Convert-ToRedactedLogMessage -Message $Message
    switch ($Level) {
        'WARN' { Write-Warning "[$stamp][$Level] $safeMessage" }
        'ERROR' { Write-Host "[$stamp][$Level] $safeMessage" -ForegroundColor Red }
        default { Write-Host "[$stamp][$Level] $safeMessage" }
    }
}

function Convert-ToRedactedLogMessage {
    param([Parameter(Mandatory = $true)][string]$Message)

    if ([string]::IsNullOrWhiteSpace($Message)) { return $Message }

    $safe = $Message -replace '(?i)\bBearer\s+[A-Za-z0-9\-_=]+\.[A-Za-z0-9\-_=]+\.[A-Za-z0-9\-_=]+', 'Bearer <redacted-token>'
    $safe = $safe -replace '(?i)"?(accessToken|refreshToken|idToken|access_token|refresh_token)"?\s*:\s*"[A-Za-z0-9\-_=]+"', '"$1":"<redacted>"'
    $safe = $safe -replace '(?i)Authorization\s*[:=]\s*([^\s,;"]+)', 'Authorization=<redacted>'
    return $safe
}

function Ensure-RunArtifactFolder {
    param([Parameter(Mandatory = $true)][string]$Path)

    $parent = Split-Path -Parent $Path
    if ([string]::IsNullOrWhiteSpace($parent)) {
        return
    }
    if (-not (Test-Path -PathType Container $parent)) {
        New-Item -Path $parent -ItemType Directory -Force | Out-Null
    }
}

function Read-SafeJsonFile {
    param([Parameter(Mandatory = $true)][string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return $null
    }
    if (-not (Test-Path -PathType Leaf $Path)) {
        return $null
    }

    try {
        $raw = Get-Content -Raw -Path $Path -ErrorAction Stop
        if ([string]::IsNullOrWhiteSpace($raw)) {
            return $null
        }
        return ConvertFrom-Json -InputObject $raw -ErrorAction Stop
    }
    catch {
        Write-Warning "Could not read JSON file '$Path': $($_.Exception.Message)"
        return $null
    }
}

function Write-SafeJsonFile {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)]$InputObject,
        [Parameter(Mandatory = $false)][int]$Depth = 20
    )

    if ([string]::IsNullOrWhiteSpace($Path) -or $null -eq $InputObject) {
        return
    }

    Ensure-RunArtifactFolder -Path $Path
    $tempPath = "$Path.tmp"
    $json = $InputObject | ConvertTo-Json -Depth $Depth
    Set-Content -Path $tempPath -Value $json -Encoding UTF8
    Move-Item -Path $tempPath -Destination $Path -Force
}

function Get-ResolvedArtifactPath {
    param(
        [Parameter(Mandatory = $true)][string]$RequestedPath,
        [Parameter(Mandatory = $true)][string]$FallbackFileName
    )

    if (-not [string]::IsNullOrWhiteSpace($RequestedPath)) {
        return $RequestedPath
    }

    if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) {
        return Join-Path $PSScriptRoot $FallbackFileName
    }

    return Join-Path (Get-Location).Path $FallbackFileName
}

function Get-TicketTemplatePath {
    param([Parameter(Mandatory = $false)][string]$TicketTemplatePath)
    return AzureSupport.TicketEngine\Get-TicketTemplatePath -TicketTemplatePath $TicketTemplatePath -RootPath $scriptRootForDefaults
}

function Get-TicketTemplate {
    param([Parameter(Mandatory = $true)][string]$Path)
    return AzureSupport.TicketEngine\Get-TicketTemplate -Path $Path
}

function Get-RequestFingerprint {
    param([Parameter(Mandatory = $true)][array]$Requests)

    if (-not $Requests -or $Requests.Count -eq 0) {
        return ""
    }

    $fingerprintInput = New-Object System.Collections.Generic.List[string]
    foreach ($request in $Requests) {
        $line = "{0}|{1}|{2}|{3}|{4}" -f $request.sub, $request.account, $request.region, $request.limit, $request.quotaType
        $fingerprintInput.Add($line.Trim())
    }

    $canonical = ($fingerprintInput | Sort-Object) -join "|"
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($canonical)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $hash = $sha.ComputeHash($bytes)
        return [System.BitConverter]::ToString($hash).Replace('-', '')
    }
    finally {
        $sha.Dispose()
    }
}

function New-RunProfileSnapshot {
    param(
        [Parameter(Mandatory = $true)][string]$TokenMode,
        [Parameter(Mandatory = $true)][int]$DelaySeconds,
        [Parameter(Mandatory = $true)][int]$RequestsPerMinute,
        [Parameter(Mandatory = $true)][int]$MaxRetries,
        [Parameter(Mandatory = $true)][int]$BaseRetrySeconds,
        [Parameter(Mandatory = $true)][bool]$RotateFingerprint,
        [Parameter(Mandatory = $true)][bool]$TryAzCliToken,
        [Parameter(Mandatory = $true)][bool]$UseDeviceCodeLogin,
        [Parameter(Mandatory = $true)][bool]$ProxyUseDefaultCredentials,
        [Parameter(Mandatory = $false)][string]$ProxyUrl,
        [Parameter(Mandatory = $false)][string[]]$ProxyPool = @(),
        [Parameter(Mandatory = $true)][bool]$DryRun,
        [Parameter(Mandatory = $true)][bool]$StopOnFirstFailure,
        [Parameter(Mandatory = $true)][int]$MaxRequests,
        [Parameter(Mandatory = $true)][bool]$RetryFailedRequests,
        [Parameter(Mandatory = $false)][bool]$ResumeFromState = $false,
        [Parameter(Mandatory = $false)][string]$CancelSignalPath,
        [Parameter(Mandatory = $false)][string]$RunStatePath
    )

    return [ordered]@{
        profileVersion = 1
        createdAt = (Get-Date).ToString('o')
        tokenSource = $TokenMode
        runSettings = @{
            DelaySeconds = $DelaySeconds
            RequestsPerMinute = $RequestsPerMinute
            MaxRetries = $MaxRetries
            BaseRetrySeconds = $BaseRetrySeconds
            RotateFingerprint = $RotateFingerprint
            MaxRequests = $MaxRequests
            TryAzCliToken = $TryAzCliToken
            UseDeviceCodeLogin = $UseDeviceCodeLogin
            StopOnFirstFailure = $StopOnFirstFailure
        }
        execution = @{
            DryRun = $DryRun
            RetryFailedRequests = $RetryFailedRequests
            ResumeFromState = $ResumeFromState
            CancelSignalPath = $CancelSignalPath
        }
        proxy = @{
            Url = $ProxyUrl
            UseDefaultCredentials = $ProxyUseDefaultCredentials
            Pool = @($ProxyPool)
        }
        resume = @{
            RunStatePath = $RunStatePath
        }
        defaults = @{
            Region = "eastus"
            QuotaType = "LowPriority"
        }
    }
}

function New-RunStateSnapshot {
    param(
        [Parameter(Mandatory = $true)][array]$Requests,
        [Parameter(Mandatory = $true)][string]$RequestFingerprint,
        [Parameter(Mandatory = $true)][bool]$StopOnFirstFailure,
        [Parameter(Mandatory = $false)][string]$RunProfilePath
    )

    $requestQueue = New-Object System.Collections.ArrayList
    $index = 0
    foreach ($request in $Requests) {
        $index++
        [void]$requestQueue.Add([pscustomobject]@{
            index = $index
            status = "Pending"
            subscription = $request.sub
            account = $request.account
            region = $request.region
            limit = $request.limit
            quotaType = $request.quotaType
            payload = $request.payload
            proxyUrl = $null
            ticket = $null
            attempts = 0
            retryCount = 0
            durationSeconds = $null
            startedAt = $null
            completedAt = $null
            error = $null
            skipReason = $null
            timelineStartUtc = $null
            timelineEndUtc = $null
            interRequestDelaySeconds = $null
            throttleWaitSeconds = $null
            timeline = $null
            attemptTimeline = @()
        })
    }

    return [ordered]@{
        runId = [guid]::NewGuid().ToString()
        status = "Running"
        requestedAction = "Run"
        createdAt = (Get-Date).ToString('o')
        updatedAt = (Get-Date).ToString('o')
        lastError = $null
        stopOnFirstFailure = $StopOnFirstFailure
        requestFingerprint = $RequestFingerprint
        runProfile = $RunProfilePath
        requestQueue = @($requestQueue)
        totalRequests = $index
        completedRequests = 0
    }
}

function Read-RunStateSnapshot {
    param([Parameter(Mandatory = $true)][string]$Path)
    return Read-SafeJsonFile -Path $Path
}

function Update-RunStateCounters {
    param($RunState)

    if ($null -eq $RunState -or -not $RunState.requestQueue) {
        return
    }

    $RunState.completedRequests = @($RunState.requestQueue | Where-Object { $_.status -in @('Submitted', 'DryRun') }).Count
}

function Write-RunStateSnapshot {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)]$RunState
    )

    if ($null -eq $RunState) {
        return
    }

    Update-RunStateCounters -RunState $RunState
    $RunState.updatedAt = (Get-Date).ToString('o')
    Write-SafeJsonFile -Path $Path -InputObject $RunState -Depth 20
}

function Get-PendingRequestStateItems {
    param(
        [Parameter(Mandatory = $true)]$RunState,
        [Parameter(Mandatory = $false)][bool]$RetryFailedRequests = $false
    )

    if ($null -eq $RunState -or -not $RunState.requestQueue) {
        return @()
    }

    if ($RetryFailedRequests) {
        return @($RunState.requestQueue | Where-Object { $_.status -eq 'Failed' } | Sort-Object index)
    }

    return @($RunState.requestQueue | Where-Object { $_.status -notin @('Submitted', 'DryRun') } | Sort-Object index)
}

function ConvertTo-ProxyPoolEntry {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Entry,
        [Parameter(Mandatory = $false)][int]$Index = 0
    )

    $trimmed = $Entry.Trim()
    if ([string]::IsNullOrWhiteSpace($trimmed)) {
        throw "Proxy entry #$Index is empty."
    }

    $scheme = "http"
    $rest = $trimmed
    if ($trimmed -match '^(?<scheme>https?)://(?<rest>.+)$') {
        $scheme = $Matches["scheme"].ToLowerInvariant()
        $rest = $Matches["rest"]
    }

    $proxyHost = $null
    $port = 0
    $username = $null
    $password = $null

    if ($rest -match '^(?<host>[^:\s]+):(?<port>\d+):(?<username>[^:\s]+):(?<password>.+)$') {
        $proxyHost = $Matches["host"]
        $port = [int]$Matches["port"]
        $username = $Matches["username"]
        $password = $Matches["password"]
    }
    elseif ($rest -match '^(?<username>[^:@\s]+):(?<password>[^@\s]+)@(?<host>[^:\s]+):(?<port>\d+)$') {
        $proxyHost = $Matches["host"]
        $port = [int]$Matches["port"]
        $username = $Matches["username"]
        $password = $Matches["password"]
    }
    elseif ($rest -match '^(?<host>[^:\s]+):(?<port>\d+)$') {
        $proxyHost = $Matches["host"]
        $port = [int]$Matches["port"]
    }
    else {
        throw "Proxy entry #$Index '$trimmed' is invalid. Use host:port:user:password, user:password@host:port, or http(s)://host:port."
    }

    if ($port -lt 1 -or $port -gt 65535) {
        throw "Proxy entry #$Index '$trimmed' has an invalid port '$port'."
    }

    $proxyUrl = "{0}://{1}:{2}" -f $scheme, $proxyHost, $port
    $credential = $null
    if (-not [string]::IsNullOrWhiteSpace($username)) {
        $securePassword = ConvertTo-SecureString -String ([string]$password) -AsPlainText -Force
        $credential = New-Object System.Management.Automation.PSCredential($username, $securePassword)
    }

    return [pscustomobject]@{
        Raw = $trimmed
        ProxyUrl = $proxyUrl
        Host = $proxyHost
        Port = $port
        Username = $username
        Credential = $credential
    }
}

function Resolve-ProxyPoolEntries {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][string]$ProxyUrl,
        [Parameter(Mandatory = $false)][string[]]$ProxyPool = @()
    )

    $candidateValues = New-Object System.Collections.Generic.List[string]
    foreach ($entry in @($ProxyPool)) {
        if ($null -eq $entry) {
            continue
        }
        $text = [string]$entry
        if ([string]::IsNullOrWhiteSpace($text)) {
            continue
        }
        $candidateValues.Add($text.Trim())
    }

    if ($candidateValues.Count -eq 0 -and -not [string]::IsNullOrWhiteSpace($ProxyUrl)) {
        $proxyUrlText = $ProxyUrl.Trim()
        $shouldSplit = $proxyUrlText -match "[,;`r`n]"
        $looksLikeCredentialProxy = $proxyUrlText -match '^[^:\s]+:\d+:[^:\s]+:.+$'

        if ($shouldSplit) {
            foreach ($item in ($proxyUrlText -split '[,;\r\n]')) {
                if (-not [string]::IsNullOrWhiteSpace($item)) {
                    $candidateValues.Add($item.Trim())
                }
            }
        }
        elseif ($looksLikeCredentialProxy) {
            $candidateValues.Add($proxyUrlText)
        }
    }

    if ($candidateValues.Count -eq 0) {
        return @()
    }

    $entries = New-Object System.Collections.Generic.List[object]
    $index = 0
    foreach ($candidate in $candidateValues) {
        $index++
        $entries.Add((ConvertTo-ProxyPoolEntry -Entry $candidate -Index $index))
    }

    return $entries.ToArray()
}

function Test-RunPreflight {
    param(
        [Parameter(Mandatory = $true)][array]$Requests,
        [Parameter(Mandatory = $false)][string]$Token,
        [Parameter(Mandatory = $true)][bool]$TryAzCliToken,
        [Parameter(Mandatory = $true)][bool]$DryRun,
        [Parameter(Mandatory = $true)][int]$DelaySeconds,
        [Parameter(Mandatory = $true)][int]$RequestsPerMinute,
        [Parameter(Mandatory = $false)][string]$ProxyUrl,
        [Parameter(Mandatory = $false)][string[]]$ProxyPool = @(),
        [Parameter(Mandatory = $true)][bool]$ProxyUseDefaultCredentials,
        [Parameter(Mandatory = $false)][pscredential]$ProxyCredential
    )

    $issues = New-Object System.Collections.Generic.List[string]

    if (-not $Requests -or $Requests.Count -eq 0) {
        $issues.Add("No quota requests were provided.")
    }

    if (-not $DryRun -and [string]::IsNullOrWhiteSpace($Token) -and -not $TryAzCliToken) {
        $issues.Add("Token is required. Set -Token, set AZURE_BEARER_TOKEN, or enable -TryAzCliToken.")
    }

    if ($DelaySeconds -lt 0) {
        $issues.Add("DelaySeconds must be 0 or greater.")
    }

    if ($RequestsPerMinute -lt 1 -or $RequestsPerMinute -gt 120) {
        $issues.Add("RequestsPerMinute must be between 1 and 120.")
    }

    if ($ProxyUseDefaultCredentials -and $ProxyCredential) {
        $issues.Add("Set either -ProxyUseDefaultCredentials or -ProxyCredential, not both.")
    }

    if ($ProxyCredential -and [string]::IsNullOrWhiteSpace($ProxyUrl)) {
        $issues.Add("ProxyCredential requires ProxyUrl.")
    }

    if (-not [string]::IsNullOrWhiteSpace($ProxyUrl) -and -not $ProxyUrl.StartsWith("http", [System.StringComparison]::OrdinalIgnoreCase)) {
        $issues.Add("ProxyUrl should be a full URL such as https://proxy:port.")
    }

    $resolvedProxyPool = @()
    try {
        $resolvedProxyPool = @(Resolve-ProxyPoolEntries -ProxyUrl $ProxyUrl -ProxyPool $ProxyPool)
    }
    catch {
        $issues.Add("Proxy pool validation failed: $($_.Exception.Message)")
    }

    if ($resolvedProxyPool.Count -gt 0) {
        if ($ProxyUseDefaultCredentials) {
            $issues.Add("ProxyUseDefaultCredentials cannot be used together with ProxyPool.")
        }
        if ($ProxyCredential) {
            $issues.Add("ProxyCredential cannot be used together with ProxyPool.")
        }
    }

    if ($Requests -and $Requests.Count -gt 0) {
        $invalidLimits = @($Requests | Where-Object { -not $_.PSObject.Properties['limit'] -or $null -eq $_.limit -or [int]$_.limit -lt 1 })
        if ($invalidLimits.Count -gt 0) {
            $issues.Add("All requests must include a positive integer limit.")
        }

        $dupes = @($Requests | Group-Object -Property { "$($_.sub)|$($_.account)|$($_.region)" } | Where-Object { $_.Count -gt 1 })
        if ($dupes.Count -gt 0) {
            $issues.Add("Duplicate request combos detected for subscription/account/region.")
        }
    }

    return $issues.ToArray()
}

function Resolve-TemplateTokens {
    param(
        [Parameter(Mandatory = $true)][string]$Template,
        [Parameter(Mandatory = $false)]$Tokens
    )

    if ([string]::IsNullOrWhiteSpace($Template)) { return $Template }
    if ($null -eq $Tokens) { return $Template }

    $result = $Template
    foreach ($token in $Tokens.GetEnumerator()) {
        $placeholder = "{" + $token.Key + "}"
        $result = $result.Replace($placeholder, [string]$token.Value)
    }
    return $result
}

function Get-EffectiveInterRequestDelaySeconds {
    param(
        [Parameter(Mandatory = $true)][int]$ConfiguredDelaySeconds,
        [Parameter(Mandatory = $true)][int]$ConfiguredRequestsPerMinute
    )

    $rateDelay = [int][math]::Ceiling(60.0 / $ConfiguredRequestsPerMinute)
    return [math]::Max($ConfiguredDelaySeconds, $rateDelay)
}

function Get-ErrorResponseBody {
    param([Parameter(Mandatory = $true)]$ErrorRecord)

    if ($ErrorRecord.ErrorDetails -and -not [string]::IsNullOrWhiteSpace($ErrorRecord.ErrorDetails.Message)) {
        return $ErrorRecord.ErrorDetails.Message
    }

    $response = Get-ExceptionResponse -ErrorRecord $ErrorRecord
    if ($null -eq $response) {
        return $null
    }

    try {
        $stream = $response.GetResponseStream()
        if ($null -eq $stream) { return $null }

        $reader = New-Object System.IO.StreamReader($stream)
        $body = $reader.ReadToEnd()
        $reader.Dispose()
        $stream.Dispose()
        return $body
    }
    catch {
        return $null
    }
}

function Get-ExceptionResponse {
    param([Parameter(Mandatory = $true)]$ErrorRecord)

    if ($null -eq $ErrorRecord -or $null -eq $ErrorRecord.Exception) {
        return $null
    }

    $responseProperty = $ErrorRecord.Exception.PSObject.Properties['Response']
    if ($null -eq $responseProperty) {
        return $null
    }

    return $responseProperty.Value
}

function Get-StatusCode {
    param([Parameter(Mandatory = $true)]$ErrorRecord)

    try {
        $response = Get-ExceptionResponse -ErrorRecord $ErrorRecord
        if ($null -eq $response) { return $null }
        return [int]$response.StatusCode
    }
    catch { return $null }
}

function Is-ThrottledResponse {
    param(
        [Parameter(Mandatory = $false)]$StatusCode,
        [Parameter(Mandatory = $false)][string]$ResponseBody,
        [Parameter(Mandatory = $false)][string]$Message
    )

    if ($StatusCode -eq 429) { return $true }

    $blob = "$ResponseBody $Message"
    if ($blob -match '(?i)too\s*many\s*requests|throttl') {
        return $true
    }

    return $false
}

function Get-RetryAfterSeconds {
    param([Parameter(Mandatory = $true)]$ErrorRecord)

    try {
        $response = Get-ExceptionResponse -ErrorRecord $ErrorRecord
        if ($null -eq $response) { return $null }
        $retryAfter = $response.Headers["Retry-After"]
        if ([string]::IsNullOrWhiteSpace($retryAfter)) { return $null }

        $secs = 0
        if ([int]::TryParse($retryAfter, [ref]$secs)) {
            return $secs
        }

        $dt = [datetime]::MinValue
        if ([datetime]::TryParse($retryAfter, [ref]$dt)) {
            $delta = [math]::Ceiling(($dt.ToUniversalTime() - (Get-Date).ToUniversalTime()).TotalSeconds)
            return [math]::Max(1, $delta)
        }

        return $null
    }
    catch { return $null }
}

function Invoke-AzCommand {
    param([Parameter(Mandatory = $true)][string[]]$Args)

    $azPath = Get-Command az -ErrorAction SilentlyContinue
    if (-not $azPath) {
        throw "Azure CLI (az) was not found in PATH."
    }

    $output = & az @Args 2>&1

    if ($azPath.CommandType -eq "Application" -and $LASTEXITCODE -ne 0) {
        throw "az $($Args -join ' ') failed: $output"
    }

    if ($output -is [array]) {
        return [string]::Join([Environment]::NewLine, $output)
    }
    return [string]$output
}

function Invoke-AzDeviceCodeLogin {
    param([Parameter(Mandatory = $false)][string]$TenantId)

    Write-Host "Running Azure device-code login..."
    Write-Host "Azure CLI will print a login URL and one-time code below."

    $loginArgs = @("login", "--use-device-code")
    if (-not [string]::IsNullOrWhiteSpace($TenantId)) {
        $loginArgs += @("--tenant", $TenantId)
        Write-Host "Tenant-scoped login requested for tenant: $TenantId"
    }

    $output = Invoke-AzCommand -Args $loginArgs
    if ($output) {
        foreach ($line in ($output -split "`r?`n")) {
            if (-not [string]::IsNullOrWhiteSpace($line)) {
                Write-Host $line
            }
        }
    }
}

function Get-AccessTokenFromAzCli {
    param(
        [Parameter(Mandatory = $false)][string]$SubscriptionId,
        [Parameter(Mandatory = $false)][string]$TenantId,
        [Parameter(Mandatory = $false)][bool]$ThrowOnError = $false
    )

    try {
        $args = @("account", "get-access-token", "--resource", "https://management.azure.com/", "-o", "json")
        if (-not [string]::IsNullOrWhiteSpace($SubscriptionId)) {
            $args += @("--subscription", $SubscriptionId)
        }
        if (-not [string]::IsNullOrWhiteSpace($TenantId)) {
            $args += @("--tenant", $TenantId)
        }

        $raw = Invoke-AzCommand -Args $args
    $tokenObj = ConvertFrom-Json -InputObject $raw
        if ($null -eq $tokenObj -or [string]::IsNullOrWhiteSpace($tokenObj.accessToken)) {
            if ($ThrowOnError) { throw "Azure CLI returned no access token." }
            return $null
        }

        return $tokenObj.accessToken
    }
    catch {
        if ($ThrowOnError) {
            throw
        }
        return $null
    }
}

function Get-SubscriptionTenantMapFromAzCli {
    param([string[]]$FilterSubscriptionIds)

    $raw = Invoke-AzCommand -Args @("account", "list", "--all", "-o", "json")
    $accounts = ConvertFrom-Json -InputObject $raw

    $map = @{}
    foreach ($acct in $accounts) {
        if ([string]::IsNullOrWhiteSpace($acct.id) -or [string]::IsNullOrWhiteSpace($acct.tenantId)) {
            continue
        }

        if ($FilterSubscriptionIds -and $FilterSubscriptionIds.Count -gt 0 -and ($FilterSubscriptionIds -notcontains $acct.id)) {
            continue
        }

        $map[$acct.id] = $acct.tenantId
    }

    return $map
}

function Get-SubscriptionTenantIdFromAzCli {
    param([Parameter(Mandatory = $true)][string]$SubscriptionId)

    try {
        $raw = Invoke-AzCommand -Args @("account", "show", "--subscription", $SubscriptionId, "-o", "json")
        $acct = ConvertFrom-Json -InputObject $raw
        if ($null -ne $acct -and -not [string]::IsNullOrWhiteSpace($acct.tenantId)) {
            return $acct.tenantId
        }
    }
    catch {
        return $null
    }

    return $null
}

function Get-TenantIdFromUnauthorizedBody {
    param([Parameter(Mandatory = $false)][string]$ResponseBody)

    if ([string]::IsNullOrWhiteSpace($ResponseBody)) {
        return $null
    }

    $mustMatch = [regex]::Match($ResponseBody, 'must match.*?sts\.windows\.net/([0-9a-fA-F-]{36})/', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [System.Text.RegularExpressions.RegexOptions]::Singleline)
    if ($mustMatch.Success) { return $mustMatch.Groups[1].Value }

    $mustMatchLogin = [regex]::Match($ResponseBody, 'authority.*?login\.windows\.net/([0-9a-fA-F-]{36})', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [System.Text.RegularExpressions.RegexOptions]::Singleline)
    if ($mustMatchLogin.Success) { return $mustMatchLogin.Groups[1].Value }

    $allMatches = [regex]::Matches($ResponseBody, 'sts\.windows\.net/([0-9a-fA-F-]{36})/')
    if ($allMatches.Count -gt 0) {
        return $allMatches[$allMatches.Count - 1].Groups[1].Value
    }

    $m = [regex]::Match($ResponseBody, 'login\.windows\.net/([0-9a-fA-F-]{36})')
    if ($m.Success) { return $m.Groups[1].Value }

    return $null
}

function Resolve-AzCliTokenForSubscription {
    param(
        [Parameter(Mandatory = $true)][string]$SubscriptionId,
        [Parameter(Mandatory = $false)][string]$KnownTenantId,
        [Parameter(Mandatory = $false)][bool]$AllowDeviceCodeLogin = $false
    )

    $candidateTenants = New-Object System.Collections.Generic.List[string]
    if (-not [string]::IsNullOrWhiteSpace($KnownTenantId)) {
        $candidateTenants.Add($KnownTenantId)
    }

    if ($candidateTenants.Count -eq 0) {
        $tenantFromAz = Get-SubscriptionTenantIdFromAzCli -SubscriptionId $SubscriptionId
        if (-not [string]::IsNullOrWhiteSpace($tenantFromAz)) {
            $candidateTenants.Add($tenantFromAz)
        }
    }

    foreach ($tenant in $candidateTenants) {
        $token = Get-AccessTokenFromAzCli -SubscriptionId $SubscriptionId -TenantId $tenant
        if (-not [string]::IsNullOrWhiteSpace($token)) {
            return @{ token = $token; tenant = $tenant }
        }

        if ($AllowDeviceCodeLogin) {
            Write-Warning "Unable to get token for subscription '$SubscriptionId' and tenant '$tenant'. Trying device-code login for that tenant."
            try {
                Invoke-AzDeviceCodeLogin -TenantId $tenant
                $token = Get-AccessTokenFromAzCli -SubscriptionId $SubscriptionId -TenantId $tenant -ThrowOnError $true
                return @{ token = $token; tenant = $tenant }
            }
            catch {
                Write-Warning "Tenant-scoped device-code login/token retrieval failed for tenant '$tenant': $($_.Exception.Message)"
            }
        }
    }

    $subToken = Get-AccessTokenFromAzCli -SubscriptionId $SubscriptionId
    if (-not [string]::IsNullOrWhiteSpace($subToken)) {
        return @{ token = $subToken; tenant = $null }
    }

    if ($AllowDeviceCodeLogin) {
        Write-Warning "Unable to acquire token for subscription '$SubscriptionId' from current Azure CLI context. Trying generic device-code login and retrying once."
        try {
            Invoke-AzDeviceCodeLogin
            $subToken = Get-AccessTokenFromAzCli -SubscriptionId $SubscriptionId -ThrowOnError $true
            if (-not [string]::IsNullOrWhiteSpace($subToken)) {
                return @{ token = $subToken; tenant = $null }
            }
        }
        catch {
            Write-Warning "Generic device-code login/token retrieval failed for subscription '$SubscriptionId': $($_.Exception.Message)"
        }
    }

    return $null
}

function Get-SubscriptionsFromAzCli {
    param([string[]]$RequestedIds)

    if ($RequestedIds -and $RequestedIds.Count -gt 0) {
        return $RequestedIds
    }

    $raw = Invoke-AzCommand -Args @("account", "list", "--query", "[].id", "-o", "tsv")
    $subs = @()
    foreach ($line in ($raw -split "`r?`n")) {
        if (-not [string]::IsNullOrWhiteSpace($line)) {
            $subs += $line.Trim()
        }
    }
    return $subs
}

function Get-BatchRequestsFromAzCli {
    param([string[]]$SubscriptionList)

    function Normalize-JsonCollection {
        param([Parameter(Mandatory = $true)]$InputObject)
        if ($null -eq $InputObject) { return @() }
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

    $discovered = @()
    foreach ($sub in $SubscriptionList) {
        $raw = Invoke-AzCommand -Args @("batch", "account", "list", "--subscription", $sub, "-o", "json")
        $parsed = ConvertFrom-Json -InputObject $raw
        $accounts = Normalize-JsonCollection -InputObject $parsed

        foreach ($a in $accounts) {
            $regionCandidates = New-Object System.Collections.Generic.List[string]
            if ($a.PSObject.Properties.Name -contains 'locations' -and $a.locations) {
                foreach ($candidate in $a.locations) {
                    if (-not [string]::IsNullOrWhiteSpace([string]$candidate)) {
                        $regionCandidates.Add([string]$candidate)
                    }
                }
            }

            if ($regionCandidates.Count -eq 0 -and -not [string]::IsNullOrWhiteSpace($a.location)) {
                $regionCandidates.Add([string]$a.location)
            }

            if ($regionCandidates.Count -eq 0) {
                $regionCandidates.Add("eastus")
            }

            $normalizedRegions = New-Object System.Collections.Generic.List[string]
            foreach ($candidate in $regionCandidates | Select-Object -Unique) {
                $normalized = "$candidate".Trim()
                if ([string]::IsNullOrWhiteSpace($normalized)) {
                    continue
                }
                if ($normalizedRegions -notcontains $normalized) {
                    $null = $normalizedRegions.Add($normalized)
                }
            }

            if ($normalizedRegions.Count -eq 0) {
                $normalizedRegions.Add("eastus")
            }

            foreach ($region in $normalizedRegions) {
                $discovered += @{
                    sub = $sub
                    account = $a.name
                    region = $region
                    regionList = @($normalizedRegions)
                }
            }
        }
    }

    return $discovered
}

function Expand-AccountRegionRequests {
    param(
        [Parameter(Mandatory = $true)][object[]]$Requests
    )

    $expanded = New-Object System.Collections.Generic.List[object]

    if (-not $Requests -or $Requests.Count -eq 0) {
        return @()
    }

    foreach ($request in $Requests) {
        if ($null -eq $request) { continue }

        $regionCandidates = @()
        $regionListValue = Resolve-RequestFieldValue -Request $request -FieldNames @('regionList')
        if ($null -ne $regionListValue) {
            $regionCandidates = @($regionListValue)
        }
        else {
            $regionValue = Resolve-RequestFieldValue -Request $request -FieldNames @('region')
            if ($null -ne $regionValue) {
                $regionCandidates = @($regionValue)
            }
        }

        $normalizedRegions = New-Object System.Collections.Generic.List[string]
        foreach ($candidate in $regionCandidates) {
            if ($null -eq $candidate) { continue }

            $splitRegions = @($candidate)
            if ($candidate -is [string]) {
                $splitRegions = $candidate -split ","
            }
            elseif (-not ($candidate -is [System.Collections.IEnumerable])) {
                $splitRegions = @($candidate.ToString())
            }

            foreach ($region in $splitRegions) {
                if ($null -eq $region) { continue }

                $normalized = "$region".Trim()
                if ([string]::IsNullOrWhiteSpace($normalized)) { continue }
                if (-not ($normalizedRegions -contains $normalized)) {
                    $null = $normalizedRegions.Add($normalized)
                }
            }
        }

        if ($normalizedRegions.Count -eq 0) {
            $null = $normalizedRegions.Add("eastus")
        }

        $baseRequest = @{}
        if ($request -is [hashtable]) {
            foreach ($pair in $request.GetEnumerator()) {
                if ($pair.Key -in @("region", "regionList")) { continue }
                $baseRequest[$pair.Key] = $pair.Value
            }
        }
        else {
            foreach ($property in $request.PSObject.Properties) {
                if ($property.Name -in @("region", "regionList")) { continue }
                $baseRequest[$property.Name] = $property.Value
            }
        }

        foreach ($region in $normalizedRegions) {
            $ticketRequest = @{}
            foreach ($pair in $baseRequest.GetEnumerator()) {
                $ticketRequest[$pair.Key] = $pair.Value
            }
            $ticketRequest["region"] = $region
            $expanded.Add([pscustomobject]$ticketRequest)
        }
    }

    return $expanded.ToArray()
}

function Resolve-RequestFieldValue {
    param(
        [Parameter(Mandatory = $true)]$Request,
        [Parameter(Mandatory = $true)][string[]]$FieldNames
    )

    foreach ($name in $FieldNames) {
        if ($Request -is [hashtable] -and $Request.ContainsKey($name)) {
            $value = $Request[$name]
            if ($null -ne $value) {
                return $value
            }
        }

        $property = $Request.PSObject.Properties[$name]
        if ($null -ne $property) {
            $value = $property.Value
            if ($null -ne $value) {
                return $value
            }
        }

        try {
            $value = $Request.$name
            if ($null -ne $value) {
                return $value
            }
        }
        catch {
            # Try next candidate field name.
        }
    }

    return $null
}

function Invoke-AzureSupportBatchQuotaRun {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [array]$Requests,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [AllowEmptyString()]
        [string]$Token = $env:AZURE_BEARER_TOKEN,

        [Parameter(Mandatory = $false)]
        [string]$TicketTemplatePath,

        [Parameter(Mandatory = $false)]
        [string]$ContactFirstName,

        [Parameter(Mandatory = $false)]
        [string]$ContactLastName,

        [Parameter(Mandatory = $false)]
        [string]$PreferredContactMethod,

        [Parameter(Mandatory = $false)]
        [string]$PrimaryEmailAddress,

        [Parameter(Mandatory = $false)]
        [string]$PreferredTimeZone,

        [Parameter(Mandatory = $false)]
        [string]$Country,

        [Parameter(Mandatory = $false)]
        [string]$PreferredSupportLanguage,

        [Parameter(Mandatory = $false)]
        [string[]]$AdditionalEmailAddresses,

        [Parameter(Mandatory = $false)]
        [string]$AcceptLanguage,

        [Parameter(Mandatory = $false)]
        [string]$ProblemClassificationId,

        [Parameter(Mandatory = $false)]
        [string]$ServiceId,

        [Parameter(Mandatory = $false)]
        [string]$Severity,

        [Parameter(Mandatory = $false)]
        [string]$Title,

        [Parameter(Mandatory = $false)]
        [string]$DescriptionTemplate,

        [Parameter(Mandatory = $false)]
        [string]$AdvancedDiagnosticConsent,

        [Parameter(Mandatory = $false)]
        [Nullable[bool]]$Require24X7Response,

        [Parameter(Mandatory = $false)]
        [string]$SupportPlanId,

        [Parameter(Mandatory = $false)]
        [string]$QuotaChangeRequestVersion,

        [Parameter(Mandatory = $false)]
        [string]$QuotaChangeRequestSubType,

        [Parameter(Mandatory = $false)]
        [string]$QuotaRequestType,

        [Parameter(Mandatory = $false)]
        [Nullable[int]]$NewLimit,

        [Parameter(Mandatory = $false)]
        [int]$DelaySeconds = 23,

        [Parameter(Mandatory = $false)]
        [int]$MaxRequests = 0,

        [Parameter(Mandatory = $false)]
        [switch]$DryRun,

        [Parameter(Mandatory = $false)]
        [string]$ProxyUrl,

        [Parameter(Mandatory = $false)]
        [switch]$ProxyUseDefaultCredentials,

        [Parameter(Mandatory = $false)]
        [pscredential]$ProxyCredential,

        [Parameter(Mandatory = $false)]
        [string[]]$ProxyPool = @(),

        [Parameter(Mandatory = $false)]
        [int]$MaxRetries = 6,

        [Parameter(Mandatory = $false)]
        [int]$BaseRetrySeconds = 25,

        [Parameter(Mandatory = $false)]
        [switch]$RotateFingerprint = $true,

        [Parameter(Mandatory = $false)]
        [bool]$TryAzCliToken = $true,

        [Parameter(Mandatory = $false)]
        [bool]$UseDeviceCodeLogin = $false,

        [Parameter(Mandatory = $false)]
        [ValidateRange(1, 120)]
        [int]$RequestsPerMinute = 2,

        [Parameter(Mandatory = $false)]
        [bool]$StopOnFirstFailure = $false,

        [Parameter(Mandatory = $false)]
        [string]$RunStatePath,

        [Parameter(Mandatory = $false)]
        [string]$RunProfilePath,

        [Parameter(Mandatory = $false)]
        [bool]$ResumeFromState = $false,

        [Parameter(Mandatory = $false)]
        [bool]$RetryFailedRequests = $false,

        [Parameter(Mandatory = $false)]
        [string]$CancelSignalPath,

        [Parameter(Mandatory = $false)]
        [string]$ResultJsonPath,

        [Parameter(Mandatory = $false)]
        [string]$ResultCsvPath
    )

    if ($UseDeviceCodeLogin -and -not $TryAzCliToken) {
        Write-Warning "-UseDeviceCodeLogin requires Azure CLI token retrieval. Enabling -TryAzCliToken automatically."
        $TryAzCliToken = $true
    }

    if ($MaxRequests -lt 0) { throw "MaxRequests cannot be negative." }
    if ($MaxRetries -lt 0) { throw "MaxRetries cannot be negative." }
    if ($BaseRetrySeconds -lt 1) { throw "BaseRetrySeconds must be >= 1." }
    if (-not $Requests -or $Requests.Count -eq 0) { throw "No quota requests were provided." }

    $effectiveContactDetails = $null
    $effectiveProblemClassificationId = $null
    $effectiveServiceId = $null
    $effectiveSeverity = $null
    $effectiveTitle = $null
    $effectiveDescriptionTemplate = $null
    $effectiveAdvancedDiagnosticConsent = $null
    $effectiveRequire24X7Response = $null
    $effectiveSupportPlanId = $null
    $effectiveQuotaChangeRequestVersion = $null
    $effectiveQuotaRequestType = $null
    $effectiveNewLimit = $null
    $effectiveQuotaChangeRequestSubType = $null
    $effectiveAcceptLanguage = $null

    $RunStatePath = Get-ResolvedArtifactPath -RequestedPath $RunStatePath -FallbackFileName "azure-support-ticket-run-state.json"

    $effectiveDelaySeconds = Get-EffectiveInterRequestDelaySeconds -ConfiguredDelaySeconds $DelaySeconds -ConfiguredRequestsPerMinute $RequestsPerMinute
    Write-Log -Message "Using effective inter-request delay of $effectiveDelaySeconds second(s) (DelaySeconds=$DelaySeconds, RequestsPerMinute=$RequestsPerMinute)."

    $normalizedRequests = @(Expand-AccountRegionRequests -Requests $Requests)
    if (-not $normalizedRequests -or $normalizedRequests.Count -eq 0) {
        throw "No valid account-region mappings were provided."
    }

    $validatedRequests = New-Object System.Collections.Generic.List[object]
    $requestIndex = 0
    foreach ($request in $normalizedRequests) {
        $requestIndex++

        $subscription = Resolve-RequestFieldValue -Request $request -FieldNames @('sub', 'subscription', 'subscriptionId', 'SubscriptionId', 'Subscription')
        $account = Resolve-RequestFieldValue -Request $request -FieldNames @('account', 'accountName', 'AccountName', 'name')
        $region = Resolve-RequestFieldValue -Request $request -FieldNames @('region', 'location')

        if ([string]::IsNullOrWhiteSpace($subscription)) {
            throw "Request #$requestIndex is missing required subscription (sub/subscription/subscriptionId)."
        }

        if ([string]::IsNullOrWhiteSpace($account)) {
            throw "Request #$requestIndex is missing required account (account/accountName/name)."
        }

        if ([string]::IsNullOrWhiteSpace($region)) {
            $region = "eastus"
        }

        $limit = $effectiveNewLimit
        $limitRaw = Resolve-RequestFieldValue -Request $request -FieldNames @('newLimit', 'NewLimit', 'limit')
        if ($null -ne $limitRaw) {
            $limitParsed = 0
            if (-not [int]::TryParse([string]$limitRaw, [ref]$limitParsed) -or $limitParsed -lt 0) {
                throw "Request #$requestIndex has an invalid newLimit '$limitRaw'."
            }

            $limit = $limitParsed
        }

        $quotaType = Resolve-RequestFieldValue -Request $request -FieldNames @('quotaType', 'type', 'Type')
        if ([string]::IsNullOrWhiteSpace($quotaType)) {
            $quotaType = $effectiveQuotaRequestType
        }

        $validatedRequests.Add([pscustomobject]@{
            Index = $requestIndex
            sub = $subscription.Trim()
            account = $account.Trim()
            region = $region.Trim()
            limit = $limit
            quotaType = $quotaType
            payload = $request
        })
    }

    if ($MaxRequests -gt 0) {
        $validatedRequests = @($validatedRequests | Select-Object -First $MaxRequests)
    }

    $preflight = Test-AzureSupportPreFlight -Requests $validatedRequests -Token $Token -TryAzCliToken $TryAzCliToken -UseDeviceCodeLogin $UseDeviceCodeLogin -DelaySeconds $DelaySeconds -MaxRequests $MaxRequests -MaxRetries $MaxRetries -BaseRetrySeconds $BaseRetrySeconds -RotateFingerprint:$RotateFingerprint -RequestsPerMinute $RequestsPerMinute -ProxyUrl $ProxyUrl -ProxyUseDefaultCredentials:$ProxyUseDefaultCredentials -ProxyCredential $ProxyCredential -RunStatePath $RunStatePath -StopOnFirstFailure $StopOnFirstFailure
    if (-not $preflight.IsValid) {
        throw ($preflight.Errors -join '; ')
    }
    foreach ($warning in @($preflight.Warnings)) {
        Write-Log -Level WARN -Message $warning
    }

    $resolvedProxyPoolEntries = @(Resolve-ProxyPoolEntries -ProxyUrl $ProxyUrl -ProxyPool $ProxyPool)
    if ($resolvedProxyPoolEntries.Count -gt 0) {
        Write-Log -Message "Proxy pool enabled. Rotating across $($resolvedProxyPoolEntries.Count) proxy endpoint(s)."
    }

    $requestFingerprint = Get-RequestFingerprint -Requests $validatedRequests
    $RunStatePath = Get-ResolvedArtifactPath -RequestedPath $RunStatePath -FallbackFileName "azure-support-ticket-run-state.json"
    $runState = $null

    if ($ResumeFromState -or $RetryFailedRequests) {
        $runState = Read-RunStateSnapshot -Path $RunStatePath
        if ($runState -and $runState.requestFingerprint -and $runState.requestFingerprint -ne $requestFingerprint) {
            Write-Warning "Request set changed since last run state was created. Starting a new persisted run state."
            $runState = $null
        }
    }

    if ($null -eq $runState) {
        $runState = New-RunStateSnapshot -Requests $validatedRequests -RequestFingerprint $requestFingerprint -StopOnFirstFailure $StopOnFirstFailure -RunProfilePath $RunProfilePath
    }
    else {
        $stateByKey = @{}
        foreach ($entry in @($runState.requestQueue)) {
            if ($null -eq $entry) {
                continue
            }
            $stateByKey["$($entry.subscription)|$($entry.account)|$($entry.region)"] = [pscustomobject]$entry
        }

        $refreshedQueue = New-Object System.Collections.Generic.List[object]
        foreach ($request in $validatedRequests) {
            $lookup = "$($request.sub)|$($request.account)|$($request.region)"
            $stateEntry = $stateByKey[$lookup]
            $stateEntry = if ($null -ne $stateEntry) { [pscustomobject]$stateEntry } else { $null }
            if ($null -eq $stateEntry) {
                $stateEntry = [pscustomobject]@{
                    index = $request.Index
                    status = "Pending"
                    subscription = $request.sub
                    account = $request.account
                    region = $request.region
                    limit = $request.limit
                    quotaType = $request.quotaType
                    payload = $request.payload
                    proxyUrl = $null
                    ticket = $null
                    attempts = 0
                    retryCount = 0
                    durationSeconds = $null
                    skipReason = $null
                    error = $null
                    startedAt = $null
                    completedAt = $null
                    timelineStartUtc = $null
                    timelineEndUtc = $null
                    interRequestDelaySeconds = $null
                    throttleWaitSeconds = $null
                    timeline = $null
                    attemptTimeline = @()
                }
            }
            else {
                $stateEntry.index = $request.Index
                $stateEntry.subscription = $request.sub
                $stateEntry.account = $request.account
                $stateEntry.region = $request.region
                $stateEntry.limit = $request.limit
                $stateEntry.quotaType = $request.quotaType
                $stateEntry.payload = $request.payload
                if (-not ($stateEntry.PSObject.Properties['proxyUrl'])) {
                    $stateEntry | Add-Member -NotePropertyName proxyUrl -NotePropertyValue $null
                }
                if (-not ($stateEntry.PSObject.Properties['retryCount'])) {
                    $stateEntry | Add-Member -NotePropertyName retryCount -NotePropertyValue 0
                }
                if (-not ($stateEntry.PSObject.Properties['timeline'])) {
                    $stateEntry | Add-Member -NotePropertyName timeline -NotePropertyValue $null
                }
                if (-not ($stateEntry.PSObject.Properties['attemptTimeline'])) {
                    $stateEntry | Add-Member -NotePropertyName attemptTimeline -NotePropertyValue @()
                }
                if (-not ($stateEntry.PSObject.Properties['timelineStartUtc'])) {
                    $stateEntry | Add-Member -NotePropertyName timelineStartUtc -NotePropertyValue $null
                }
                if (-not ($stateEntry.PSObject.Properties['timelineEndUtc'])) {
                    $stateEntry | Add-Member -NotePropertyName timelineEndUtc -NotePropertyValue $null
                }
                if (-not ($stateEntry.PSObject.Properties['interRequestDelaySeconds'])) {
                    $stateEntry | Add-Member -NotePropertyName interRequestDelaySeconds -NotePropertyValue $null
                }
                if (-not ($stateEntry.PSObject.Properties['throttleWaitSeconds'])) {
                    $stateEntry | Add-Member -NotePropertyName throttleWaitSeconds -NotePropertyValue $null
                }
            }
            $refreshedQueue.Add($stateEntry) | Out-Null
        }

        $runState.requestQueue = [object[]]$refreshedQueue
        $runState.requestFingerprint = $requestFingerprint
        if (-not [string]::IsNullOrWhiteSpace($RunProfilePath)) {
            $runState.runProfile = $RunProfilePath
        }
        $runState.status = "Running"
        $runState.lastError = $null
        $runState.updatedAt = (Get-Date).ToString('o')
    }

    if ($ResumeFromState -or $RetryFailedRequests) {
        $pendingStateItems = Get-PendingRequestStateItems -RunState $runState -RetryFailedRequests:$RetryFailedRequests
    }
    else {
        $pendingStateItems = @($runState.requestQueue)
    }

    Write-RunStateSnapshot -Path $RunStatePath -RunState $runState

    $tokenFromAzCli = $false
    $workingToken = $Token
    if ([string]::IsNullOrWhiteSpace($workingToken) -and $TryAzCliToken) {
        if ($UseDeviceCodeLogin) {
            try {
                Invoke-AzDeviceCodeLogin
            }
            catch {
                Write-Warning "Device-code login failed: $($_.Exception.Message)"
            }
        }

        $workingToken = Get-AccessTokenFromAzCli
        if (-not [string]::IsNullOrWhiteSpace($workingToken)) {
            $tokenFromAzCli = $true
            Write-Host "Using access token from Azure CLI."
        }
    }

    if ([string]::IsNullOrWhiteSpace($workingToken)) {
        Write-Warning "Token is required. Pass -Token, set AZURE_BEARER_TOKEN, or enable Azure CLI token retrieval."
        Write-Host "Example: ./create_azure_support_tickets.ps1 -TryAzCliToken `$true -UseDeviceCodeLogin"
        throw "Token was not resolved."
    }

    $normalizedToken = $workingToken.Trim()
    if ($normalizedToken -match '^[Bb]earer\s+') {
        $normalizedToken = $normalizedToken -replace '^[Bb]earer\s+', ''
    }

    if ($null -eq $effectiveContactDetails -or $null -eq $effectiveProblemClassificationId -or $null -eq $effectiveQuotaRequestType -or $null -eq $effectiveNewLimit) {
        $resolvedTicketTemplatePath = Get-TicketTemplatePath -TicketTemplatePath $TicketTemplatePath
        $ticketTemplate = Get-TicketTemplate -Path $resolvedTicketTemplatePath

        if (-not $ticketTemplate.contactDetails) {
            throw "Ticket template '$resolvedTicketTemplatePath' is missing required contactDetails."
        }

        $templateContact = $ticketTemplate.contactDetails
        if ($null -eq $effectiveContactDetails) {
            $additionalEmailSource = if ($PSBoundParameters.ContainsKey("AdditionalEmailAddresses")) {
                @($AdditionalEmailAddresses)
            }
            else {
                @($templateContact.additionalEmailAddresses)
            }
            $resolvedAdditionalEmailAddresses = New-Object 'System.Collections.Generic.List[string]'
            foreach ($email in @($additionalEmailSource)) {
                if ($null -eq $email) {
                    continue
                }
                $text = [string]$email
                if ([string]::IsNullOrWhiteSpace($text)) {
                    continue
                }
                $null = $resolvedAdditionalEmailAddresses.Add($text.Trim())
            }

            $effectiveContactDetails = @{
                firstName = if ($PSBoundParameters.ContainsKey("ContactFirstName")) { $ContactFirstName } else { [string]$templateContact.firstName }
                lastName = if ($PSBoundParameters.ContainsKey("ContactLastName")) { $ContactLastName } else { [string]$templateContact.lastName }
                preferredContactMethod = if ($PSBoundParameters.ContainsKey("PreferredContactMethod")) { $PreferredContactMethod } else { [string]$templateContact.preferredContactMethod }
                primaryEmailAddress = if ($PSBoundParameters.ContainsKey("PrimaryEmailAddress")) { $PrimaryEmailAddress } else { [string]$templateContact.primaryEmailAddress }
                preferredTimeZone = if ($PSBoundParameters.ContainsKey("PreferredTimeZone")) { $PreferredTimeZone } else { [string]$templateContact.preferredTimeZone }
                country = if ($PSBoundParameters.ContainsKey("Country")) { $Country } else { [string]$templateContact.country }
                preferredSupportLanguage = if ($PSBoundParameters.ContainsKey("PreferredSupportLanguage")) { $PreferredSupportLanguage } else { [string]$templateContact.preferredSupportLanguage }
                additionalEmailAddresses = $resolvedAdditionalEmailAddresses.ToArray()
            }
        }

        if ($null -eq $effectiveAcceptLanguage) { $effectiveAcceptLanguage = if ($PSBoundParameters.ContainsKey("AcceptLanguage")) { $AcceptLanguage } else { [string]$ticketTemplate.acceptLanguage } }
        if ($null -eq $effectiveProblemClassificationId) { $effectiveProblemClassificationId = if ($PSBoundParameters.ContainsKey("ProblemClassificationId")) { $ProblemClassificationId } else { [string]$ticketTemplate.problemClassificationId } }
        if ($null -eq $effectiveServiceId) { $effectiveServiceId = if ($PSBoundParameters.ContainsKey("ServiceId")) { $ServiceId } else { [string]$ticketTemplate.serviceId } }
        if ($null -eq $effectiveSeverity) { $effectiveSeverity = if ($PSBoundParameters.ContainsKey("Severity")) { $Severity } else { [string]$ticketTemplate.severity } }
        if ($null -eq $effectiveTitle) { $effectiveTitle = if ($PSBoundParameters.ContainsKey("Title")) { $Title } else { [string]$ticketTemplate.title } }
        if ($null -eq $effectiveDescriptionTemplate) { $effectiveDescriptionTemplate = if ($PSBoundParameters.ContainsKey("DescriptionTemplate")) { $DescriptionTemplate } else { [string]$ticketTemplate.descriptionTemplate } }
        if ($null -eq $effectiveAdvancedDiagnosticConsent) { $effectiveAdvancedDiagnosticConsent = if ($PSBoundParameters.ContainsKey("AdvancedDiagnosticConsent")) { $AdvancedDiagnosticConsent } else { [string]$ticketTemplate.advancedDiagnosticConsent } }
        if ($null -eq $effectiveRequire24X7Response) { $effectiveRequire24X7Response = if ($PSBoundParameters.ContainsKey("Require24X7Response")) { [bool]$Require24X7Response } else { [bool]$ticketTemplate.require24X7Response } }
        if ($null -eq $effectiveSupportPlanId) { $effectiveSupportPlanId = if ($PSBoundParameters.ContainsKey("SupportPlanId")) { $SupportPlanId } else { [string]$ticketTemplate.supportPlanId } }
        if ($null -eq $effectiveQuotaChangeRequestVersion) { $effectiveQuotaChangeRequestVersion = if ($PSBoundParameters.ContainsKey("QuotaChangeRequestVersion")) { $QuotaChangeRequestVersion } else { [string]$ticketTemplate.quotaChangeRequestVersion } }
        if ($null -eq $effectiveQuotaChangeRequestSubType) { $effectiveQuotaChangeRequestSubType = if ($PSBoundParameters.ContainsKey("QuotaChangeRequestSubType")) { $QuotaChangeRequestSubType } else { [string]$ticketTemplate.quotaChangeRequestSubType } }
        if ($null -eq $effectiveQuotaRequestType) { $effectiveQuotaRequestType = if ($PSBoundParameters.ContainsKey("QuotaRequestType")) { $QuotaRequestType } else { [string]$ticketTemplate.quotaRequestType } }
        if ($null -eq $effectiveNewLimit) { $effectiveNewLimit = if ($PSBoundParameters.ContainsKey("NewLimit") -and $null -ne $NewLimit) { [int]$NewLimit } else { [int]$ticketTemplate.newLimit } }

        if ($effectiveNewLimit -le 0) {
            throw "The resolved NewLimit must be a positive integer. Update the template or pass -NewLimit."
        }
    }

    $baseHeaders = @{
        Accept = "*/*"
        "Accept-Language" = if ([string]::IsNullOrWhiteSpace($effectiveAcceptLanguage)) { "en" } else { $effectiveAcceptLanguage }
        "Content-Type" = "application/json"
    }

    $subscriptionTenantMap = @{}
    if ($tokenFromAzCli) {
        try {
            $subscriptionTenantMap = Get-SubscriptionTenantMapFromAzCli -FilterSubscriptionIds ($validatedRequests | ForEach-Object { $_.sub } | Select-Object -Unique)
            if ($subscriptionTenantMap.Count -gt 0) {
                Write-Host "Resolved tenant IDs for $($subscriptionTenantMap.Count) subscriptions."
            }
        }
        catch {
            Write-Warning "Unable to resolve subscription -> tenant mapping from Azure CLI: $($_.Exception.Message)"
        }
    }

    $results = New-Object System.Collections.Generic.List[object]

:RequestRunLoop
    foreach ($stateEntry in @($pendingStateItems | Sort-Object index)) {
        if (-not [string]::IsNullOrWhiteSpace($CancelSignalPath) -and (Test-Path $CancelSignalPath)) {
            $runState.status = "Cancelled"
            $runState.lastError = "Execution cancelled by signal file '$CancelSignalPath'."
            Write-Log -Level WARN -Message $runState.lastError
            break RequestRunLoop
        }

        $r = $validatedRequests | Where-Object { $_.Index -eq [int]$stateEntry.index } | Select-Object -First 1
        if ($null -eq $r) {
            continue
        }

        $stateEntry.startedAt = (Get-Date).ToUniversalTime().ToString('o')
        $stateEntry.status = 'Running'
        $stateEntry.skipReason = $null
        $stateEntry.error = $null
        $stateEntry.ticket = $null
        Write-RunStateSnapshot -Path $RunStatePath -RunState $runState

        $requestStart = Get-Date
        $requestStartUtc = $requestStart.ToUniversalTime().ToString('o')
        $succeeded = $false
        $failureMessage = $null
        $attemptCount = 0
        $throttleWaitSeconds = 0
        $attemptTimeline = New-Object System.Collections.Generic.List[object]
        $throttledResponseDetected = $false

        $requestProxyUrl = $ProxyUrl
        $requestProxyUseDefaultCredentials = [bool]$ProxyUseDefaultCredentials
        $requestProxyCredential = $ProxyCredential
        if ($resolvedProxyPoolEntries.Count -gt 0) {
            $proxyPoolIndex = ([int]$stateEntry.index - 1) % $resolvedProxyPoolEntries.Count
            if ($proxyPoolIndex -lt 0) {
                $proxyPoolIndex = 0
            }
            $proxyEntry = $resolvedProxyPoolEntries[$proxyPoolIndex]
            $requestProxyUrl = $proxyEntry.ProxyUrl
            $requestProxyUseDefaultCredentials = $false
            $requestProxyCredential = $proxyEntry.Credential
        }
        $stateEntry.proxyUrl = $requestProxyUrl

        $ticket = [guid]::NewGuid().ToString()
        $payload = (@{
            AccountName = $r.account
            NewLimit = $r.limit
            Type = $r.quotaType
        } | ConvertTo-Json -Compress)

        $description = Resolve-TemplateTokens -Template $effectiveDescriptionTemplate -Tokens @{
            region = $r.region
            newLimit = $r.limit
            quotaType = $r.quotaType
        }

        $bodyObject = @{
            properties = @{
                contactDetails = $effectiveContactDetails
                description = $description
                problemClassificationId = $effectiveProblemClassificationId
                serviceId = $effectiveServiceId
                severity = $effectiveSeverity
                title = $effectiveTitle
                advancedDiagnosticConsent = $effectiveAdvancedDiagnosticConsent
                require24X7Response = $effectiveRequire24X7Response
                supportPlanId = $effectiveSupportPlanId
                quotaTicketDetails = @{
                    quotaChangeRequestVersion = $effectiveQuotaChangeRequestVersion
                    quotaChangeRequestSubType = $effectiveQuotaChangeRequestSubType
                    quotaChangeRequests = @(
                        @{
                            region = $r.region
                            payload = $payload
                        }
                    )
                }
            }
        }

        $body = $bodyObject | ConvertTo-Json -Depth 10 -Compress
        $url = "https://management.azure.com/subscriptions/$($r.sub)/providers/Microsoft.Support/supportTickets/${ticket}?api-version=2025-06-01-preview"

        if ($DryRun) {
            $dryRunEnd = Get-Date
            $dryRunDuration = [math]::Round(($dryRunEnd - $requestStart).TotalSeconds, 2)
            $attemptTimeline.Add([pscustomobject]@{
                attempt = 1
                status = 'DryRun'
                startedAtUtc = $requestStartUtc
                endedAtUtc = $dryRunEnd.ToUniversalTime().ToString('o')
                durationSeconds = [math]::Round(($dryRunEnd - $requestStart).TotalSeconds, 3)
                statusCode = 0
                retryAfterSeconds = 0
                sleepSeconds = 0
                error = $null
                reason = 'Dry run: request not sent'
            })
            Write-Host "[DRY RUN] Prepared quota request -> $($r.account)"
            Write-Log -Message "[DRY RUN] Would submit request for $($r.account) in $($r.region) using ticket $ticket."
            $stateEntry.status = 'DryRun'
            $stateEntry.ticket = $ticket
            $stateEntry.attempts = 1
            $stateEntry.durationSeconds = $dryRunDuration
            $stateEntry.retryCount = 0
            $stateEntry.timelineStartUtc = $requestStartUtc
            $stateEntry.timelineEndUtc = $dryRunEnd.ToUniversalTime().ToString('o')
            $stateEntry.interRequestDelaySeconds = $effectiveDelaySeconds
            $stateEntry.throttleWaitSeconds = 0
            $stateEntry.timeline = @{
                startUtc = $requestStartUtc
                endUtc = $dryRunEnd.ToUniversalTime().ToString('o')
                durationSeconds = $dryRunDuration
                interRequestDelaySeconds = $effectiveDelaySeconds
                maxRetries = $MaxRetries
                attempts = $attemptTimeline
                throttleWaitSeconds = 0
            }
            $stateEntry.attemptTimeline = $attemptTimeline
            $stateEntry.completedAt = $dryRunEnd.ToUniversalTime().ToString('o')
            $stateEntry.error = $null
            Write-RunStateSnapshot -Path $RunStatePath -RunState $runState
            $results.Add([pscustomobject]@{
                requestIndex = $r.Index
                account = $r.account
                subscription = $r.sub
                region = $r.region
                proxyUrl = $requestProxyUrl
                ticket = $ticket
                status = 'DryRun'
                attempts = 1
                durationSeconds = $dryRunDuration
                retryCount = 0
                timelineStartUtc = $requestStartUtc
                timelineEndUtc = $dryRunEnd.ToUniversalTime().ToString('o')
                interRequestDelaySeconds = $effectiveDelaySeconds
                throttleWaitSeconds = 0
                timeline = @{
                    startUtc = $requestStartUtc
                    endUtc = $dryRunEnd.ToUniversalTime().ToString('o')
                    durationSeconds = $dryRunDuration
                    interRequestDelaySeconds = $effectiveDelaySeconds
                    maxRetries = $MaxRetries
                    attempts = $attemptTimeline
                    throttleWaitSeconds = 0
                }
                attemptTimeline = $attemptTimeline
                error = $null
            })
            if ($effectiveDelaySeconds -gt 0) {
                Start-Sleep -Seconds $effectiveDelaySeconds
            }
            continue
        }

        for ($attempt = 0; $attempt -le $MaxRetries; $attempt++) {
            $attemptCount = $attempt + 1
            $attemptStart = Get-Date
            $attemptStartUtc = $attemptStart.ToUniversalTime().ToString('o')
            $attemptSleepSeconds = 0
            $attemptStatus = 'Failed'
            $attemptError = $null
            $attemptStatusCode = $null
            $requestHeaders = @{}
            foreach ($k in $baseHeaders.Keys) { $requestHeaders[$k] = $baseHeaders[$k] }

            if ($tokenFromAzCli) {
                $tenantForSub = $null
                if ($subscriptionTenantMap.ContainsKey($r.sub)) {
                    $tenantForSub = $subscriptionTenantMap[$r.sub]
                }

                $tokenResolution = Resolve-AzCliTokenForSubscription -SubscriptionId $r.sub -KnownTenantId $tenantForSub -AllowDeviceCodeLogin:$UseDeviceCodeLogin
                if ($null -eq $tokenResolution -or [string]::IsNullOrWhiteSpace($tokenResolution.token)) {
                    $failureMessage = "Unable to acquire an Azure CLI token for subscription '$($r.sub)'. Ensure your account has access to that subscription tenant or run with -UseDeviceCodeLogin:`$true and authenticate for the required tenant."
                    $attemptTimeline.Add([pscustomobject]@{
                        attempt = 1
                        status = 'TokenAcquisitionFailed'
                        startedAtUtc = (Get-Date).ToUniversalTime().ToString('o')
                        endedAtUtc = (Get-Date).ToUniversalTime().ToString('o')
                        durationSeconds = 0
                        statusCode = 0
                        retryAfterSeconds = 0
                        sleepSeconds = 0
                        error = $failureMessage
                        reason = 'Unable to obtain token for subscription'
                    })
                    break
                }

                if (-not [string]::IsNullOrWhiteSpace($tokenResolution.tenant)) {
                    $subscriptionTenantMap[$r.sub] = $tokenResolution.tenant
                }

                $normalizedToken = $tokenResolution.token.Trim()
            }

            $requestHeaders["Authorization"] = "Bearer $normalizedToken"

            if ($RotateFingerprint) {
                $requestHeaders["User-Agent"] = "AzureQuotaBot/1.0 fp-$([guid]::NewGuid().ToString('N'))"
                $requestHeaders["x-ms-client-request-id"] = [guid]::NewGuid().ToString()
                $requestHeaders["x-ms-correlation-request-id"] = [guid]::NewGuid().ToString()
            }

            $invokeParams = @{
                Method = "Put"
                Uri = $url
                Headers = $requestHeaders
                Body = [System.Text.Encoding]::UTF8.GetBytes($body)
                ContentType = "application/json; charset=utf-8"
                ErrorAction = "Stop"
            }

            if (-not [string]::IsNullOrWhiteSpace($requestProxyUrl)) {
                $invokeParams["Proxy"] = $requestProxyUrl
                if ($requestProxyUseDefaultCredentials) {
                    $invokeParams["ProxyUseDefaultCredentials"] = $true
                }
                elseif ($requestProxyCredential) {
                    $invokeParams["ProxyCredential"] = $requestProxyCredential
                }
            }

            try {
                $null = Invoke-RestMethod @invokeParams
                $attemptStatus = 'Success'
                Write-Host "Submitted quota request -> $($r.account)"
                $succeeded = $true
                $attemptTimeline.Add([pscustomobject]@{
                    attempt = $attemptCount
                    status = $attemptStatus
                    startedAtUtc = $attemptStartUtc
                    endedAtUtc = (Get-Date).ToUniversalTime().ToString('o')
                    durationSeconds = [math]::Round(((Get-Date) - $attemptStart).TotalSeconds, 3)
                    statusCode = 0
                    retryAfterSeconds = 0
                    sleepSeconds = 0
                    error = $null
                    reason = 'Success'
                })
                break
            }
            catch {
                $status = Get-StatusCode -ErrorRecord $_
                $responseBody = Get-ErrorResponseBody -ErrorRecord $_
                $redactedResponseBody = if ([string]::IsNullOrWhiteSpace($responseBody)) { '' } else { Convert-ToRedactedLogMessage -Message $responseBody }
                $attemptStatusCode = $status

                $tenantMismatch = $false
                if ($status -eq 401) {
                    $tenantMismatch = $true
                }
                if ($responseBody -and ($responseBody -match '(?i)InvalidAuthenticationTokenTenant|wrong issuer|must match the tenant')) {
                    $tenantMismatch = $true
                }

                if ($tokenFromAzCli -and $tenantMismatch -and $attempt -lt $MaxRetries) {
                    $tenantFromBody = Get-TenantIdFromUnauthorizedBody -ResponseBody $responseBody
                    if (-not [string]::IsNullOrWhiteSpace($tenantFromBody)) {
                        $subscriptionTenantMap[$r.sub] = $tenantFromBody
                        $tenantResolution = Resolve-AzCliTokenForSubscription -SubscriptionId $r.sub -KnownTenantId $tenantFromBody -AllowDeviceCodeLogin:$UseDeviceCodeLogin
                        if ($tenantResolution -and -not [string]::IsNullOrWhiteSpace($tenantResolution.token)) {
                            $normalizedToken = $tenantResolution.token.Trim()
                            Write-Warning "401 tenant mismatch for $($r.sub). Refreshed token for tenant $tenantFromBody and retrying."
                            $attemptStatus = 'RetryingAfterTenantRefresh'
                            $attemptTimeline.Add([pscustomobject]@{
                                attempt = $attemptCount
                                status = $attemptStatus
                                startedAtUtc = $attemptStartUtc
                                endedAtUtc = (Get-Date).ToUniversalTime().ToString('o')
                                durationSeconds = [math]::Round(((Get-Date) - $attemptStart).TotalSeconds, 3)
                                statusCode = $status
                                retryAfterSeconds = 0
                                sleepSeconds = 0
                                error = $null
                                reason = "tenant mismatch for subscription '$($r.sub)', refreshed token"
                            })
                            continue
                        }
                    }
                }

                $isThrottledResponse = Is-ThrottledResponse -StatusCode $status -ResponseBody $responseBody -Message $_.Exception.Message
                if ($isThrottledResponse) {
                    $throttledResponseDetected = $true
                }

                if ($isThrottledResponse -and $attempt -lt $MaxRetries) {
                    $retryAfter = Get-RetryAfterSeconds -ErrorRecord $_
                    $backoff = [math]::Min(300, $BaseRetrySeconds * [math]::Pow(2, $attempt))
                    $jitter = Get-Random -Minimum 0 -Maximum 12
                    $sleepSeconds = if ($retryAfter) { [math]::Max($retryAfter, $backoff + $jitter) } else { $backoff + $jitter }
                    $attemptSleepSeconds = $sleepSeconds
                    $throttleWaitSeconds += $sleepSeconds
                    $attemptStatus = 'ThrottledRetryScheduled'

                    Write-Warning "429 throttled for $($r.account). Retry $($attempt + 1)/$MaxRetries in ${sleepSeconds}s."
                    $attemptTimeline.Add([pscustomobject]@{
                        attempt = $attemptCount
                        status = $attemptStatus
                        startedAtUtc = $attemptStartUtc
                        endedAtUtc = (Get-Date).ToUniversalTime().ToString('o')
                        durationSeconds = [math]::Round(((Get-Date) - $attemptStart).TotalSeconds, 3)
                        statusCode = $status
                        retryAfterSeconds = $retryAfter
                        sleepSeconds = $sleepSeconds
                        error = $null
                        reason = 'throttled; retry scheduled'
                    })

                    Start-Sleep -Seconds $sleepSeconds
                    continue
                }

                $attemptStatus = 'Failed'
                if (-not [string]::IsNullOrWhiteSpace($responseBody)) {
                    $failureMessage = "REST request failed for account '$($r.account)' in subscription '$($r.sub)'. HTTP $status. $($_.Exception.Message) ResponseBody: $redactedResponseBody"
                    $attemptError = $failureMessage
                }
                else {
                    $failureMessage = "REST request failed for account '$($r.account)' in subscription '$($r.sub)'. HTTP $status. $($_.Exception.Message)"
                    $attemptError = $failureMessage
                }

                $attemptTimeline.Add([pscustomobject]@{
                    attempt = $attemptCount
                    status = $attemptStatus
                    startedAtUtc = $attemptStartUtc
                    endedAtUtc = (Get-Date).ToUniversalTime().ToString('o')
                    durationSeconds = [math]::Round(((Get-Date) - $attemptStart).TotalSeconds, 3)
                    statusCode = $attemptStatusCode
                    retryAfterSeconds = 0
                    sleepSeconds = $attemptSleepSeconds
                    error = $attemptError
                    reason = 'request failed'
                })
                break
            }
        }

        $requestEnd = Get-Date
        $requestDuration = [math]::Round(($requestEnd - $requestStart).TotalSeconds, 2)
        $requestStatus = 'Failed'
        if ($succeeded) {
            $requestStatus = 'Submitted'
        }
        $stateEntry.status = $requestStatus
        $stateEntry.ticket = $ticket
        $stateEntry.attempts = $attemptCount
        $stateEntry.durationSeconds = $requestDuration
        $stateEntry.retryCount = [Math]::Max(0, $attemptCount - 1)
        $stateEntry.timelineStartUtc = $requestStartUtc
        $stateEntry.timelineEndUtc = $requestEnd.ToUniversalTime().ToString('o')
        $stateEntry.interRequestDelaySeconds = $effectiveDelaySeconds
        $stateEntry.throttleWaitSeconds = [math]::Round($throttleWaitSeconds, 2)
        $attemptTimelineEntries = $attemptTimeline
        $timelineSnapshot = [pscustomobject]@{
            startUtc = $requestStartUtc
            endUtc = $requestEnd.ToUniversalTime().ToString('o')
            durationSeconds = $requestDuration
            interRequestDelaySeconds = $effectiveDelaySeconds
            maxRetries = $MaxRetries
            retryCount = [Math]::Max(0, $attemptCount - 1)
            throttleWaitSeconds = [math]::Round($throttleWaitSeconds, 2)
            attempts = $attemptTimelineEntries
        }
        try {
            $stateEntry.timeline = $timelineSnapshot
        }
        catch {
            # Some deserialized state entries can reject complex object assignment.
            $stateEntry.timeline = $null
        }
        $stateEntry.attemptTimeline = $attemptTimelineEntries
        $stateEntry.error = $failureMessage
        $stateEntry.completedAt = $requestEnd.ToUniversalTime().ToString('o')
        Write-RunStateSnapshot -Path $RunStatePath -RunState $runState

        $results.Add([pscustomobject]@{
            requestIndex = $r.Index
            account = $r.account
            subscription = $r.sub
            region = $r.region
            proxyUrl = $requestProxyUrl
            ticket = $ticket
            status = $requestStatus
            attempts = $attemptCount
            durationSeconds = $requestDuration
            retryCount = [Math]::Max(0, $attemptCount - 1)
            timelineStartUtc = $requestStartUtc
            timelineEndUtc = $requestEnd.ToUniversalTime().ToString('o')
            interRequestDelaySeconds = $effectiveDelaySeconds
            throttleWaitSeconds = [math]::Round($throttleWaitSeconds, 2)
            timeline = $timelineSnapshot
            attemptTimeline = $attemptTimelineEntries
            error = $failureMessage
        })

        if (-not $succeeded) {
            Write-Log -Level ERROR -Message $failureMessage
            $runState.lastError = $failureMessage
            Write-RunStateSnapshot -Path $RunStatePath -RunState $runState
            if ($throttledResponseDetected -and -not $DryRun) {
                $runState.status = "Throttled"
                Write-RunStateSnapshot -Path $RunStatePath -RunState $runState
                throw "Stopping run because account '$($r.account)' was throttled (HTTP 429)."
            }
            if ($StopOnFirstFailure) {
                $runState.status = "Failed"
                Write-RunStateSnapshot -Path $RunStatePath -RunState $runState
                throw "Stopping early because -StopOnFirstFailure is enabled."
            }
        }

        if ($effectiveDelaySeconds -gt 0) {
            Start-Sleep -Seconds $effectiveDelaySeconds
        }
    }

    if ($runState.status -ne 'Cancelled' -and $runState.status -ne 'Failed') {
        $runState.status = 'Completed'
    }
    Write-RunStateSnapshot -Path $RunStatePath -RunState $runState

    # Backfill results for any state entries that were skipped (resume/cancel) but have
    # completed status from a previous run, so the final result set is always complete.
    $resultKeys = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($r in $results) {
        [void]$resultKeys.Add("$($r.subscription)|$($r.account)|$($r.region)")
    }
    foreach ($entry in @($runState.requestQueue | Sort-Object index)) {
        if (-not $entry.status -or $entry.status -eq 'Pending' -or $entry.status -eq 'Running') {
            continue
        }
        $key = "$($entry.subscription)|$($entry.account)|$($entry.region)"
        if ($resultKeys.Contains($key)) {
            continue
        }
        $results.Add([pscustomobject]@{
            requestIndex = $entry.index
            account = $entry.account
            subscription = $entry.subscription
            region = $entry.region
            proxyUrl = if ($entry.PSObject.Properties['proxyUrl']) { $entry.proxyUrl } else { $null }
            ticket = $entry.ticket
            status = $entry.status
            attempts = if ($null -ne $entry.attempts) { $entry.attempts } else { 0 }
            durationSeconds = if ($null -ne $entry.durationSeconds) { $entry.durationSeconds } else { 0 }
            retryCount = if ($null -ne $entry.retryCount) { $entry.retryCount } else { 0 }
            timelineStartUtc = $entry.timelineStartUtc
            timelineEndUtc = $entry.timelineEndUtc
            interRequestDelaySeconds = $entry.interRequestDelaySeconds
            throttleWaitSeconds = $entry.throttleWaitSeconds
            timeline = $entry.timeline
            attemptTimeline = $entry.attemptTimeline
            error = $entry.error
        })
    }

    if (-not [string]::IsNullOrWhiteSpace($ResultJsonPath)) {
        Ensure-RunArtifactFolder -Path $ResultJsonPath
        $results | ConvertTo-Json -Depth 8 | Set-Content -Path $ResultJsonPath -Encoding UTF8
        Write-Log -Message "Saved request results to $ResultJsonPath"
    }

    if (-not [string]::IsNullOrWhiteSpace($ResultCsvPath)) {
        $resultsForCsv = foreach ($entry in $results) {
            $timelineJson = if ($entry.timeline) { ConvertTo-Json -InputObject $entry.timeline -Depth 8 -Compress } else { "" }
            $attemptTimelineJson = if ($entry.attemptTimeline) { ConvertTo-Json -InputObject $entry.attemptTimeline -Depth 8 -Compress } else { "[]" }
            [pscustomobject]@{
                requestIndex = $entry.requestIndex
                account = $entry.account
                subscription = $entry.subscription
                region = $entry.region
                proxyUrl = $entry.proxyUrl
                ticket = $entry.ticket
                status = $entry.status
                attempts = $entry.attempts
                retryCount = $entry.retryCount
                durationSeconds = $entry.durationSeconds
                timelineStartUtc = $entry.timelineStartUtc
                timelineEndUtc = $entry.timelineEndUtc
                interRequestDelaySeconds = $entry.interRequestDelaySeconds
                throttleWaitSeconds = $entry.throttleWaitSeconds
                error = $entry.error
                timelineJson = $timelineJson
                attemptTimelineJson = $attemptTimelineJson
            }
        }
        Ensure-RunArtifactFolder -Path $ResultCsvPath
        $resultsForCsv | Export-Csv -Path $ResultCsvPath -NoTypeInformation -Encoding UTF8
        Write-Log -Message "Saved request results CSV to $ResultCsvPath"
    }

    return $results
}

function Invoke-AzureSupportBatchQuotaRunQueued {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][array]$Requests,
        [Parameter(Mandatory = $false)][string]$Token = $env:AZURE_BEARER_TOKEN,
        [Parameter(Mandatory = $false)][int]$DelaySeconds = 23,
        [Parameter(Mandatory = $false)][int]$MaxRequests = 0,
        [Parameter(Mandatory = $false)][switch]$DryRun,
        [Parameter(Mandatory = $false)][string]$ProxyUrl,
        [Parameter(Mandatory = $false)][switch]$ProxyUseDefaultCredentials,
        [Parameter(Mandatory = $false)][pscredential]$ProxyCredential,
        [Parameter(Mandatory = $false)][string[]]$ProxyPool = @(),
        [Parameter(Mandatory = $false)][int]$MaxRetries = 6,
        [Parameter(Mandatory = $false)][int]$BaseRetrySeconds = 25,
        [Parameter(Mandatory = $false)][switch]$RotateFingerprint = $true,
        [Parameter(Mandatory = $false)][bool]$TryAzCliToken = $true,
        [Parameter(Mandatory = $false)][bool]$UseDeviceCodeLogin = $false,
        [Parameter(Mandatory = $false)][ValidateRange(1, 120)][int]$RequestsPerMinute = 2,
        [Parameter(Mandatory = $false)][bool]$StopOnFirstFailure = $false,
        [Parameter(Mandatory = $false)][string]$RunStatePath,
        [Parameter(Mandatory = $false)][switch]$ResumeFromState,
        [Parameter(Mandatory = $false)][switch]$RetryFailedRequests,
        [Parameter(Mandatory = $false)][string]$RunProfilePath
    )

    if (-not $Requests -or $Requests.Count -eq 0) {
        throw "No requests were prepared for execution."
    }

    $runStatePath = Get-ResolvedArtifactPath -RequestedPath $RunStatePath -FallbackFileName "azure-support-ticket-run-state.json"
    $requestFingerprint = Get-RequestFingerprint -Requests $Requests
    $runState = Read-RunStateSnapshot -Path $runStatePath

    if (
        $ResumeFromState -and
        $null -ne $runState -and
        $runState.requestFingerprint -eq $requestFingerprint -and
        $runState.requestQueue
    ) {
        $runState.status = "Running"
        $runState.requestedAction = "Run"
        Write-Host "Resuming run from existing run state: $runStatePath"
    }
    else {
        $runState = New-RunStateSnapshot -Requests $Requests -RequestFingerprint $requestFingerprint -StopOnFirstFailure $StopOnFirstFailure -RunProfilePath $RunProfilePath
    }

    $pendingItems = Get-PendingRequestStateItems -RunState $runState -RetryFailedRequests:$RetryFailedRequests
    if (-not $pendingItems -or $pendingItems.Count -eq 0) {
        $runState.status = "Completed"
        Write-RunStateSnapshot -Path $runStatePath -RunState $runState
        $finalResults = New-Object System.Collections.Generic.List[object]
        foreach ($entry in ($runState.requestQueue | Sort-Object index)) {
            $finalResults.Add([pscustomobject]@{
                account = $entry.account
                subscription = $entry.subscription
                region = $entry.region
                ticket = $entry.ticket
                status = $entry.status
                attempts = [int]$entry.attempts
                durationSeconds = $entry.durationSeconds
                error = $entry.error
            })
        }
        return $finalResults
    }

    foreach ($entry in $pendingItems) {
        $controlState = Read-RunStateSnapshot -Path $runStatePath
        if ($null -ne $controlState -and ($controlState.requestedAction -eq "Cancel" -or $controlState.status -eq "CancelRequested")) {
            $runState.status = "Cancelled"
            $runState.lastError = "Run was cancelled before next request."
            Write-RunStateSnapshot -Path $runStatePath -RunState $runState
            break
        }

        $entry.status = "Running"
        $entry.startedAt = (Get-Date).ToString("o")
        $entry.error = $null
        $entry.skipReason = $null
        $entry.attempts = 0
        $entry.durationSeconds = $null

        $singleRequest = [pscustomobject]@{
            sub = $entry.subscription
            account = $entry.account
            region = $entry.region
            newLimit = $entry.limit
            quotaType = $entry.quotaType
        }

        try {
            $runTemplatePath = Get-TicketTemplatePath -TicketTemplatePath $TicketTemplatePath
            $runParams = @{
                Requests = @($singleRequest)
                Token = $Token
                TicketTemplatePath = $runTemplatePath
                ContactFirstName = $ContactFirstName
                ContactLastName = $ContactLastName
                PreferredContactMethod = $PreferredContactMethod
                PrimaryEmailAddress = $PrimaryEmailAddress
                PreferredTimeZone = $PreferredTimeZone
                Country = $Country
                PreferredSupportLanguage = $PreferredSupportLanguage
                AdditionalEmailAddresses = $AdditionalEmailAddresses
                AcceptLanguage = $AcceptLanguage
                ProblemClassificationId = $ProblemClassificationId
                ServiceId = $ServiceId
                Severity = $Severity
                Title = $Title
                DescriptionTemplate = $DescriptionTemplate
                AdvancedDiagnosticConsent = $AdvancedDiagnosticConsent
                Require24X7Response = $Require24X7Response
                SupportPlanId = $SupportPlanId
                QuotaChangeRequestVersion = $QuotaChangeRequestVersion
                QuotaChangeRequestSubType = $QuotaChangeRequestSubType
                QuotaRequestType = $QuotaRequestType
                NewLimit = $NewLimit
                DelaySeconds = $DelaySeconds
                MaxRequests = 0
                DryRun = [bool]$DryRun
                ProxyUrl = $ProxyUrl
                ProxyUseDefaultCredentials = [bool]$ProxyUseDefaultCredentials
                ProxyCredential = $ProxyCredential
                ProxyPool = $ProxyPool
                MaxRetries = $MaxRetries
                BaseRetrySeconds = $BaseRetrySeconds
                RotateFingerprint = [bool]$RotateFingerprint
                TryAzCliToken = [bool]$TryAzCliToken
                UseDeviceCodeLogin = [bool]$UseDeviceCodeLogin
                RequestsPerMinute = $RequestsPerMinute
                StopOnFirstFailure = [bool]$StopOnFirstFailure
                RunProfilePath = $runState.runProfile
            }

            $cleanRunParams = Remove-EmptyParameters -Parameters $runParams
            $result = Invoke-AzureSupportBatchQuotaRun @cleanRunParams
        }
        catch {
            $entry.status = "Failed"
            $entry.error = $_.Exception.Message
            $entry.completedAt = (Get-Date).ToString("o")
            $entry.durationSeconds = [math]::Round(((Get-Date) - ([datetime]$entry.startedAt)).TotalSeconds, 2)
            $runState.lastError = $entry.error
            Write-RunStateSnapshot -Path $runStatePath -RunState $runState
            $isThrottleStop = (-not $DryRun) -and ($entry.error -match '(?i)\b429\b|throttl')
            if ($isThrottleStop) {
                $runState.status = "Throttled"
                Write-RunStateSnapshot -Path $runStatePath -RunState $runState
                throw
            }
            if (-not $DryRun -and $StopOnFirstFailure) {
                $runState.status = "StoppedOnFailure"
                Write-RunStateSnapshot -Path $runStatePath -RunState $runState
                throw
            }
            continue
        }

        $single = $result[0]
        $entry.status = $single.status
        $entry.ticket = $single.ticket
        $entry.attempts = [int]$single.attempts
        $entry.durationSeconds = $single.durationSeconds
        $entry.completedAt = (Get-Date).ToString("o")
        $entry.error = $single.error
        Write-RunStateSnapshot -Path $runStatePath -RunState $runState

        if ($DelaySeconds -gt 0 -and -not $DryRun -and $entry -ne $pendingItems[-1]) {
            Start-Sleep -Seconds $DelaySeconds
        }
    }

    if ($runState.status -ne "Cancelled" -and $runState.status -ne "StoppedOnFailure") {
        $runState.status = "Completed"
    }
    Write-RunStateSnapshot -Path $runStatePath -RunState $runState

    $finalResults = New-Object System.Collections.Generic.List[object]
    foreach ($entry in ($runState.requestQueue | Sort-Object index)) {
        $finalResults.Add([pscustomobject]@{
            account = $entry.account
            subscription = $entry.subscription
            region = $entry.region
            ticket = $entry.ticket
            status = $entry.status
            attempts = [int]$entry.attempts
            durationSeconds = $entry.durationSeconds
            error = $entry.error
        })
    }

    return $finalResults
}

if ($MyInvocation.InvocationName -ne '.') {
    $resolvedTicketTemplatePath = Get-TicketTemplatePath -TicketTemplatePath $TicketTemplatePath
    $ticketTemplate = Get-TicketTemplate -Path $resolvedTicketTemplatePath
    if (-not $ticketTemplate.contactDetails) {
        throw "Ticket template '$resolvedTicketTemplatePath' is missing required contactDetails."
    }

    $templateContact = $ticketTemplate.contactDetails
    $additionalEmailSource = if ($PSBoundParameters.ContainsKey("AdditionalEmailAddresses")) {
        @($AdditionalEmailAddresses)
    }
    else {
        @($templateContact.additionalEmailAddresses)
    }
    $resolvedAdditionalEmailAddresses = New-Object 'System.Collections.Generic.List[string]'
    foreach ($email in @($additionalEmailSource)) {
        if ($null -eq $email) {
            continue
        }
        $text = [string]$email
        if ([string]::IsNullOrWhiteSpace($text)) {
            continue
        }
        $null = $resolvedAdditionalEmailAddresses.Add($text.Trim())
    }

    $effectiveContactDetails = @{
        firstName = if ($PSBoundParameters.ContainsKey("ContactFirstName")) { $ContactFirstName } else { [string]$templateContact.firstName }
        lastName = if ($PSBoundParameters.ContainsKey("ContactLastName")) { $ContactLastName } else { [string]$templateContact.lastName }
        preferredContactMethod = if ($PSBoundParameters.ContainsKey("PreferredContactMethod")) { $PreferredContactMethod } else { [string]$templateContact.preferredContactMethod }
        primaryEmailAddress = if ($PSBoundParameters.ContainsKey("PrimaryEmailAddress")) { $PrimaryEmailAddress } else { [string]$templateContact.primaryEmailAddress }
        preferredTimeZone = if ($PSBoundParameters.ContainsKey("PreferredTimeZone")) { $PreferredTimeZone } else { [string]$templateContact.preferredTimeZone }
        country = if ($PSBoundParameters.ContainsKey("Country")) { $Country } else { [string]$templateContact.country }
        preferredSupportLanguage = if ($PSBoundParameters.ContainsKey("PreferredSupportLanguage")) { $PreferredSupportLanguage } else { [string]$templateContact.preferredSupportLanguage }
        additionalEmailAddresses = $resolvedAdditionalEmailAddresses.ToArray()
    }

    $effectiveAcceptLanguage = if ($PSBoundParameters.ContainsKey("AcceptLanguage")) { $AcceptLanguage } else { [string]$ticketTemplate.acceptLanguage }
    $effectiveProblemClassificationId = if ($PSBoundParameters.ContainsKey("ProblemClassificationId")) { $ProblemClassificationId } else { [string]$ticketTemplate.problemClassificationId }
    $effectiveServiceId = if ($PSBoundParameters.ContainsKey("ServiceId")) { $ServiceId } else { [string]$ticketTemplate.serviceId }
    $effectiveSeverity = if ($PSBoundParameters.ContainsKey("Severity")) { $Severity } else { [string]$ticketTemplate.severity }
    $effectiveTitle = if ($PSBoundParameters.ContainsKey("Title")) { $Title } else { [string]$ticketTemplate.title }
    $effectiveDescriptionTemplate = if ($PSBoundParameters.ContainsKey("DescriptionTemplate")) { $DescriptionTemplate } else { [string]$ticketTemplate.descriptionTemplate }
    $effectiveAdvancedDiagnosticConsent = if ($PSBoundParameters.ContainsKey("AdvancedDiagnosticConsent")) { $AdvancedDiagnosticConsent } else { [string]$ticketTemplate.advancedDiagnosticConsent }
    $effectiveRequire24X7Response = if ($PSBoundParameters.ContainsKey("Require24X7Response")) { [bool]$Require24X7Response } else { [bool]$ticketTemplate.require24X7Response }
    $effectiveSupportPlanId = if ($PSBoundParameters.ContainsKey("SupportPlanId")) { $SupportPlanId } else { [string]$ticketTemplate.supportPlanId }
    $effectiveQuotaChangeRequestVersion = if ($PSBoundParameters.ContainsKey("QuotaChangeRequestVersion")) { $QuotaChangeRequestVersion } else { [string]$ticketTemplate.quotaChangeRequestVersion }
    $effectiveQuotaChangeRequestSubType = if ($PSBoundParameters.ContainsKey("QuotaChangeRequestSubType")) { $QuotaChangeRequestSubType } else { [string]$ticketTemplate.quotaChangeRequestSubType }
    $effectiveQuotaRequestType = if ($PSBoundParameters.ContainsKey("QuotaRequestType")) { $QuotaRequestType } else { [string]$ticketTemplate.quotaRequestType }
    $effectiveNewLimit = if ($PSBoundParameters.ContainsKey("NewLimit") -and $null -ne $NewLimit) { [int]$NewLimit } else { [int]$ticketTemplate.newLimit }

    if ($effectiveNewLimit -le 0) {
        throw "The resolved NewLimit must be a positive integer. Update the template or pass -NewLimit."
    }

    $requests = @()
    if ($ticketTemplate.defaultRequests) {
        foreach ($requestTemplate in @($ticketTemplate.defaultRequests)) {
            if ($null -eq $requestTemplate) { continue }

            $templateSub = [string]$requestTemplate.sub
            $templateAccount = [string]$requestTemplate.account
            $templateRegion = [string]$requestTemplate.region

            if ([string]::IsNullOrWhiteSpace($templateSub) -or [string]::IsNullOrWhiteSpace($templateAccount) -or [string]::IsNullOrWhiteSpace($templateRegion)) {
                continue
            }

            $requests += @{
                sub = $templateSub
                account = $templateAccount
                region = $templateRegion
            }
        }
    }

    if (-not $AutoDiscoverRequests -and (-not $requests -or $requests.Count -eq 0)) {
        throw "No default requests are defined in '$resolvedTicketTemplatePath'. Either define defaultRequests in the template or run with -AutoDiscoverRequests."
    }

    if ($AutoDiscoverRequests) {
        $subs = Get-SubscriptionsFromAzCli -RequestedIds $SubscriptionIds
        $requests = Get-BatchRequestsFromAzCli -SubscriptionList $subs

        if (-not $requests -or $requests.Count -eq 0) {
            throw "No Batch accounts were discovered from Azure CLI."
        }

        Write-Host "Discovered $($requests.Count) Batch accounts from Azure."
    }

    $resolvedProfilePath = Get-ResolvedArtifactPath -RequestedPath $RunProfilePath -FallbackFileName "azure-ticket-run-profile.json"
    $resolvedStatePath = Get-ResolvedArtifactPath -RequestedPath $RunStatePath -FallbackFileName "azure-support-ticket-run-state.json"
    $boundParams = $PSBoundParameters

    if ($LoadRunProfile) {
        $loadedProfile = Read-SafeJsonFile -Path $resolvedProfilePath
        if ($null -ne $loadedProfile) {
            $loadedRunSettings = $loadedProfile.runSettings
            $loadedProxy = $loadedProfile.proxy
            $loadedExecution = $loadedProfile.execution
            if ($null -eq $loadedRunSettings) { $loadedRunSettings = $loadedProfile }
            if ($null -eq $loadedExecution) { $loadedExecution = $null }

            if (-not $boundParams.ContainsKey('DelaySeconds') -and $loadedRunSettings.DelaySeconds) {
                $DelaySeconds = [int]$loadedRunSettings.DelaySeconds
            }
            if (-not $boundParams.ContainsKey('RequestsPerMinute') -and $loadedRunSettings.RequestsPerMinute) {
                $RequestsPerMinute = [int]$loadedRunSettings.RequestsPerMinute
            }
            if (-not $boundParams.ContainsKey('MaxRetries') -and $loadedRunSettings.MaxRetries) {
                $MaxRetries = [int]$loadedRunSettings.MaxRetries
            }
            if (-not $boundParams.ContainsKey('BaseRetrySeconds') -and $loadedRunSettings.BaseRetrySeconds) {
                $BaseRetrySeconds = [int]$loadedRunSettings.BaseRetrySeconds
            }
            if (-not $boundParams.ContainsKey('RotateFingerprint') -and $loadedRunSettings.RotateFingerprint -ne $null) {
                $RotateFingerprint = $loadedRunSettings.RotateFingerprint
            }
            if (-not $boundParams.ContainsKey('MaxRequests') -and $loadedRunSettings.MaxRequests -ne $null) {
                $MaxRequests = [int]$loadedRunSettings.MaxRequests
            }
            if (-not $boundParams.ContainsKey('TryAzCliToken') -and $loadedRunSettings.TryAzCliToken -ne $null) {
                $TryAzCliToken = $loadedRunSettings.TryAzCliToken
            }
            if (-not $boundParams.ContainsKey('UseDeviceCodeLogin') -and $loadedRunSettings.UseDeviceCodeLogin -ne $null) {
                $UseDeviceCodeLogin = $loadedRunSettings.UseDeviceCodeLogin
            }
            if (-not $boundParams.ContainsKey('StopOnFirstFailure') -and $loadedRunSettings.StopOnFirstFailure -ne $null) {
                $StopOnFirstFailure = $loadedRunSettings.StopOnFirstFailure
            }
            if (-not $boundParams.ContainsKey('DryRun') -and $null -ne $loadedExecution -and $loadedExecution.DryRun -ne $null) {
                $DryRun = $loadedExecution.DryRun
            }
            if (-not $boundParams.ContainsKey('RetryFailedRequests') -and $loadedExecution.RetryFailedRequests -ne $null) {
                $RetryFailedRequests = $loadedExecution.RetryFailedRequests
            }
            if (-not $boundParams.ContainsKey('ResumeFromState') -and $loadedExecution.ResumeFromState -ne $null) {
                $ResumeFromState = $loadedExecution.ResumeFromState
            }
            if (-not $boundParams.ContainsKey('CancelSignalPath') -and $loadedExecution.CancelSignalPath) {
                $CancelSignalPath = [string]$loadedExecution.CancelSignalPath
            }
            if (-not $boundParams.ContainsKey('ProxyUrl') -and $loadedProxy.Url) {
                $ProxyUrl = [string]$loadedProxy.Url
            }
            if (-not $boundParams.ContainsKey('ProxyUseDefaultCredentials') -and $loadedProxy.UseDefaultCredentials -ne $null) {
                $ProxyUseDefaultCredentials = $loadedProxy.UseDefaultCredentials
            }
            if (-not $boundParams.ContainsKey('ProxyPool') -and $loadedProxy.Pool -ne $null) {
                $ProxyPool = @($loadedProxy.Pool)
            }
        }
    }

    $RotateFingerprint = Convert-ToBoolValue -Value $RotateFingerprint -Default $true -Name "RotateFingerprint"
    $TryAzCliToken = Convert-ToBoolValue -Value $TryAzCliToken -Default $true -Name "TryAzCliToken"
    $UseDeviceCodeLogin = Convert-ToBoolValue -Value $UseDeviceCodeLogin -Default $false -Name "UseDeviceCodeLogin"
    $StopOnFirstFailure = Convert-ToBoolValue -Value $StopOnFirstFailure -Default $false -Name "StopOnFirstFailure"
    $DryRun = Convert-ToBoolValue -Value $DryRun -Default $false -Name "DryRun"
    $ResumeFromState = Convert-ToBoolValue -Value $ResumeFromState -Default $false -Name "ResumeFromState"
    $RetryFailedRequests = Convert-ToBoolValue -Value $RetryFailedRequests -Default $false -Name "RetryFailedRequests"
    $ProxyUseDefaultCredentials = Convert-ToBoolValue -Value $ProxyUseDefaultCredentials -Default $false -Name "ProxyUseDefaultCredentials"
    if ($null -eq $ProxyPool) {
        $ProxyPool = @()
    }
    else {
        $ProxyPool = @(
            @($ProxyPool) |
                ForEach-Object { if ($null -eq $_) { '' } else { [string]$_ } } |
                ForEach-Object { $_.Trim() } |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        )
    }

    $requests = Expand-AccountRegionRequests -Requests @($requests)

    $preparedRequests = New-Object System.Collections.Generic.List[object]
    $rawIndex = 0
    foreach ($request in $requests) {
        $rawIndex++
        $subscription = Resolve-RequestFieldValue -Request $request -FieldNames @('sub', 'subscription', 'subscriptionId', 'SubscriptionId', 'Subscription')
        $account = Resolve-RequestFieldValue -Request $request -FieldNames @('account', 'accountName', 'AccountName', 'name')
        $region = Resolve-RequestFieldValue -Request $request -FieldNames @('region', 'location')
        if ([string]::IsNullOrWhiteSpace($region)) {
            $region = "eastus"
        }

        if ([string]::IsNullOrWhiteSpace($subscription)) {
            throw "Request #$rawIndex is missing required subscription (sub/subscription/subscriptionId)."
        }
        if ([string]::IsNullOrWhiteSpace($account)) {
            throw "Request #$rawIndex is missing required account (account/accountName/name)."
        }

        $limit = $effectiveNewLimit
        $limitRaw = Resolve-RequestFieldValue -Request $request -FieldNames @('newLimit', 'NewLimit', 'limit')
        if ($null -ne $limitRaw) {
            $limitParsed = 0
            if (-not [int]::TryParse([string]$limitRaw, [ref]$limitParsed) -or $limitParsed -lt 0) {
                throw "Request #$rawIndex has an invalid newLimit '$limitRaw'."
            }
            $limit = $limitParsed
        }

        $quotaType = Resolve-RequestFieldValue -Request $request -FieldNames @('quotaType', 'type', 'Type')
        if ([string]::IsNullOrWhiteSpace($quotaType)) {
            $quotaType = $effectiveQuotaRequestType
        }

        $preparedRequests.Add([pscustomobject]@{
            sub = $subscription.Trim()
            account = $account.Trim()
            region = $region.Trim()
            limit = $limit
            quotaType = $quotaType
        })
    }

    $preflight = Test-RunPreflight -Requests $preparedRequests -Token $Token -TryAzCliToken $TryAzCliToken -DryRun:$DryRun -DelaySeconds $DelaySeconds -RequestsPerMinute $RequestsPerMinute -ProxyUrl $ProxyUrl -ProxyPool $ProxyPool -ProxyUseDefaultCredentials:$ProxyUseDefaultCredentials -ProxyCredential $ProxyCredential
    if ($preflight -is [System.Management.Automation.PSCustomObject]) {
        if (-not $preflight.IsValid) {
            throw (($preflight.Errors) -join "; ")
        }
    }
    else {
        $preflightIssues = @($preflight)
        if ($preflightIssues.Count -gt 0) {
            throw ($preflightIssues -join "; ")
        }
    }

    $scriptStart = Get-Date
    $resolvedTicketTemplatePath = Get-TicketTemplatePath -TicketTemplatePath $TicketTemplatePath
    $runParams = @{
        Requests = $preparedRequests
        Token = $Token
        TicketTemplatePath = $resolvedTicketTemplatePath
        ContactFirstName = $ContactFirstName
        ContactLastName = $ContactLastName
        PreferredContactMethod = $PreferredContactMethod
        PrimaryEmailAddress = $PrimaryEmailAddress
        PreferredTimeZone = $PreferredTimeZone
        Country = $Country
        PreferredSupportLanguage = $PreferredSupportLanguage
        AdditionalEmailAddresses = $AdditionalEmailAddresses
        AcceptLanguage = $AcceptLanguage
        ProblemClassificationId = $ProblemClassificationId
        ServiceId = $ServiceId
        Severity = $Severity
        Title = $Title
        DescriptionTemplate = $DescriptionTemplate
        AdvancedDiagnosticConsent = $AdvancedDiagnosticConsent
        Require24X7Response = $Require24X7Response
        SupportPlanId = $SupportPlanId
        QuotaChangeRequestVersion = $QuotaChangeRequestVersion
        QuotaChangeRequestSubType = $QuotaChangeRequestSubType
        QuotaRequestType = $QuotaRequestType
        NewLimit = $NewLimit
        DelaySeconds = $DelaySeconds
        MaxRequests = $MaxRequests
        DryRun = [bool]$DryRun
        ProxyUrl = $ProxyUrl
        ProxyUseDefaultCredentials = [bool]$ProxyUseDefaultCredentials
        ProxyCredential = $ProxyCredential
        ProxyPool = $ProxyPool
        MaxRetries = $MaxRetries
        BaseRetrySeconds = $BaseRetrySeconds
        RotateFingerprint = [bool]$RotateFingerprint
        TryAzCliToken = [bool]$TryAzCliToken
        UseDeviceCodeLogin = [bool]$UseDeviceCodeLogin
        RequestsPerMinute = $RequestsPerMinute
        StopOnFirstFailure = [bool]$StopOnFirstFailure
        RunProfilePath = $resolvedProfilePath
        RunStatePath = $resolvedStatePath
        ResumeFromState = [bool]$ResumeFromState
        RetryFailedRequests = [bool]$RetryFailedRequests
        CancelSignalPath = $CancelSignalPath
        ResultJsonPath = $ResultJsonPath
        ResultCsvPath = $ResultCsvPath
    }
    $cleanedRunParams = Remove-EmptyParameters -Parameters $runParams
    $results = @(Invoke-AzureSupportBatchQuotaRun @cleanedRunParams)

    $submittedCount = @($results | Where-Object { $_.status -eq 'Submitted' }).Count
    $failedCount = @($results | Where-Object { $_.status -eq 'Failed' }).Count
    $dryRunCount = @($results | Where-Object { $_.status -eq 'DryRun' }).Count
    $elapsedSeconds = [math]::Round(((Get-Date) - $scriptStart).TotalSeconds, 2)

    Write-Log -Message "Run completed. Submitted=$submittedCount Failed=$failedCount DryRun=$dryRunCount Total=$($results.Count) Duration=${elapsedSeconds}s"

    if ($SaveRunProfile) {
        if (-not [string]::IsNullOrWhiteSpace($Token)) {
            $tokenSource = "Token"
        }
        elseif ($TryAzCliToken) {
            $tokenSource = "AzureCli"
        }
        else {
            $tokenSource = "None"
        }

        $runProfile = New-RunProfileSnapshot `
            -TokenMode $tokenSource `
            -DelaySeconds $DelaySeconds `
            -RequestsPerMinute $RequestsPerMinute `
            -MaxRetries $MaxRetries `
            -BaseRetrySeconds $BaseRetrySeconds `
            -RotateFingerprint $RotateFingerprint `
            -TryAzCliToken $TryAzCliToken `
            -UseDeviceCodeLogin $UseDeviceCodeLogin `
            -ProxyUseDefaultCredentials $ProxyUseDefaultCredentials `
            -ProxyUrl $ProxyUrl `
            -ProxyPool $ProxyPool `
            -DryRun:$DryRun `
            -StopOnFirstFailure $StopOnFirstFailure `
            -MaxRequests $MaxRequests `
            -RetryFailedRequests $RetryFailedRequests `
            -ResumeFromState $ResumeFromState `
            -CancelSignalPath $CancelSignalPath `
            -RunStatePath $resolvedStatePath

        Write-SafeJsonFile -Path $resolvedProfilePath -InputObject $runProfile -Depth 20
        Write-Log -Message "Saved run profile to $resolvedProfilePath"
    }


    if ($failedCount -gt 0 -and -not $DryRun) {
        exit 1
    }
}
