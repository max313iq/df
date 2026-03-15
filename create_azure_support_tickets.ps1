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
    [switch]$StopOnFirstFailure
)


Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$engineModulePath = Join-Path $scriptRoot "Modules\AzureSupport.TicketEngine.psm1"
if (-not (Test-Path -LiteralPath $engineModulePath)) {
    throw "Ticket engine module not found at '$engineModulePath'."
}
Import-Module -Name $engineModulePath -Force -ErrorAction Stop | Out-Null

$boundParameters = @{}
foreach ($entry in $PSBoundParameters.GetEnumerator()) {
    $boundParameters[$entry.Key] = $entry.Value
}

AzureSupport.TicketEngine\Invoke-AzureSupportTicketCli -BoundParameters $boundParameters -ScriptRootOverride $scriptRoot
