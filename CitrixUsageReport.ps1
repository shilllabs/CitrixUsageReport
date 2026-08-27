#requires -Version 5.1
# ============================================================================
#  Citrix Usage Report - GENERATED FILE, DO NOT EDIT DIRECTLY
#  Built : 2026-08-24T19:09:39Z
#  Source: https://github.com/shilllabs/CitrixUsageReport/
#  Edit the files under src/ and re-run build.ps1 instead.
# ============================================================================

# ----------------------------------------------------------------------------
# region 00-Header.ps1
# ----------------------------------------------------------------------------
<#
.SYNOPSIS
    Citrix Usage Report - reports unique users and concurrent session usage
    from a Citrix environment over the requested periods.

.DESCRIPTION
    Connects to the Citrix Monitor Service OData API on a Citrix Virtual Apps
    and Desktops (on-premises) or Citrix DaaS (cloud) environment, collects
    session usage data, and produces a self-contained branded HTML report
    covering 30, 60, and 90 day windows.

    The script is strictly read-only. It issues HTTP GET requests against the
    Monitor OData endpoint and, for cloud environments, one POST to obtain a
    bearer token. It changes nothing in the Citrix environment.

    On-premises, point this at a Delivery Controller -- the machine running
    the Citrix Monitor Service. Citrix Director is a separate IIS web
    application that queries the Monitor Service; on smaller deployments it
    may be installed on the same machine as a Delivery Controller, but a
    standalone Director web server does not run the Monitor Service itself
    and will not work.

    Run with no parameters to get an interactive dialog.

.PARAMETER DeliveryController
    Hostname or FQDN of the Citrix Delivery Controller (the machine running
    the Citrix Monitor Service), for on-premises environments. On smaller
    deployments Director may be installed on the same machine, but a
    standalone Director web server will not work -- point this at the
    Delivery Controller. Example: ddc01.contoso.com

.PARAMETER Protocol
    Which scheme to use for the Monitor OData endpoint: Https (default) or
    Http. The Monitor Service can be published over either -- it is a setting
    on the Delivery Controller's IIS site, not something this tool can infer.
    Your Citrix administrator will know which applies. HTTPS is the right
    answer wherever it is available; select Http only if the Monitor Service
    on this Delivery Controller is not published over HTTPS. Running over
    HTTP transmits session data (usernames, machine names, delivery group
    names) unencrypted -- the report will say so plainly if you choose it.
    Ignored for cloud environments, which are always HTTPS. An explicit
    http:// or https:// typed into -DeliveryController overrides this.

.PARAMETER Environment
    Which Citrix environment to target. Determines the API endpoints used.
    OnPremises        - a Delivery Controller you name
    CloudCommercial   - Citrix Cloud, US / EU / Asia Pacific South
    CloudJapan        - Citrix Cloud, Japan region
    CloudGovernment   - Citrix Cloud Government (US Gov)

    The cloud region is never inferred. Selecting the wrong region would send
    credentials to the wrong sovereign endpoint, so it is always explicit.

.PARAMETER CustomerId
    Citrix Cloud Customer ID. Required for all cloud environments.

.PARAMETER ClientId
    Citrix Cloud API client ID. Required for all cloud environments.

.PARAMETER ClientSecret
    Citrix Cloud API client secret, as a SecureString.

.PARAMETER Credential
    Windows credential for on-premises authentication. Omit to use the
    identity of the account running the script.

.PARAMETER Days
    Reporting windows in days. Defaults to 30, 60 and 90. Comma-separate
    multiple values with no spaces: -Days 30,60,90

.PARAMETER OutputPath
    Directory to write the report and exports into.

.PARAMETER IncludeDeliveryGroups
    Break results down per delivery group.

.PARAMETER IncludeApplications
    Break results down by session type and published application. Requires
    fetching ApplicationInstances, which is slower on large sites.

.PARAMETER IncludeClientDevices
    Break results down by client device, address and Workspace app version.
    Requires fetching Connections, which is slower on large sites.

.PARAMETER IncludeTrend
    Include the daily unique-user and daily peak-concurrency trend.

.PARAMETER Anonymize
    Replace usernames with stable User-0001 identifiers in the report and
    exports. Writes a separate identity-map.csv that stays on this machine.

.PARAMETER ExportRawData
    Write data.json and CSV exports alongside the HTML report.

.PARAMETER DemoData
    Generate a synthetic dataset instead of contacting Citrix. Use this to
    preview the report without an environment.

.PARAMETER NoGui
    Skip the graphical dialog and use console prompts instead.

.EXAMPLE
    .\CitrixUsageReport.ps1

    Opens the interactive dialog.

.EXAMPLE
    .\CitrixUsageReport.ps1 -DeliveryController ddc01.contoso.com -OutputPath C:\Reports

    Runs non-interactively against an on-premises Delivery Controller as the
    current user, over HTTPS.

.EXAMPLE
    .\CitrixUsageReport.ps1 -DeliveryController ddc01.contoso.com -Protocol Http -OutputPath C:\Reports

    Same, but against a Delivery Controller whose Monitor Service is
    published over HTTP rather than HTTPS. The report will carry a visible
    warning that the run was performed unencrypted.

.EXAMPLE
    .\CitrixUsageReport.ps1 -DemoData -OutputPath C:\Temp

    Produces a sample report from synthetic data, with no Citrix environment.

.NOTES
    Data availability depends on the site's Citrix license edition. Raw session
    data is groomed at 90 days on Premium, 31 days on Advanced, and 7 days on
    other editions. The script detects the real limit and labels any window it
    could not fully cover.
#>
[CmdletBinding(DefaultParameterSetName = 'Interactive')]
param(
    [Parameter(ParameterSetName = 'OnPremises')]
    [string] $DeliveryController,

    [Parameter(ParameterSetName = 'OnPremises')]
    [ValidateSet('Https', 'Http')]
    [string] $Protocol = 'Https',

    [ValidateSet('OnPremises', 'CloudCommercial', 'CloudJapan', 'CloudGovernment')]
    [string] $Environment,

    [Parameter(ParameterSetName = 'Cloud')]
    [string] $CustomerId,

    [Parameter(ParameterSetName = 'Cloud')]
    [string] $ClientId,

    [Parameter(ParameterSetName = 'Cloud')]
    [System.Security.SecureString] $ClientSecret,

    [Parameter(ParameterSetName = 'OnPremises')]
    [System.Management.Automation.PSCredential] $Credential,

    # [string[]], not [int[]] with [ValidateRange]: under `-File` -- the
    # invocation the runbook must recommend, since it is the one that works
    # with -ExecutionPolicy Bypass against a Restricted default policy --
    # PowerShell hands the script raw argv strings and parses none of it as
    # PowerShell syntax. `-Days 30,60,90` therefore arrives as the single
    # string "30,60,90", not an array, and an [int[]] parameter coerces that
    # single string to an int by stripping the commas as thousands
    # separators (306090), which then fails ValidateRange. Kept as
    # unvalidated strings here and parsed with ConvertTo-DaysArray
    # (src/90-Gui.ps1) inside Invoke-CitrixUsageAudit, the only place both
    # the real argv shape and that function are available together -- see
    # the note there for why the parsing cannot happen here instead.
    [string[]] $Days = @('30', '60', '90'),

    [string] $OutputPath = (Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'CitrixUsageReport'),

    [switch] $IncludeDeliveryGroups,
    [switch] $IncludeApplications,
    [switch] $IncludeClientDevices,
    [switch] $IncludeTrend,
    [switch] $Anonymize,
    [switch] $ExportRawData,
    [switch] $DemoData,
    [switch] $NoGui
)

# endregion 00-Header.ps1

# ----------------------------------------------------------------------------
# region 10-Logging.ps1
# ----------------------------------------------------------------------------
# ============================================================================
#  Logging
#
#  Every message the script emits passes through here, which makes this the
#  single place responsible for ensuring no credential or bearer token is ever
#  written to the console or to the log file on disk.
# ============================================================================

# Literal secret values registered at runtime (client secrets, tokens).
$script:AuditSecrets = New-Object System.Collections.Generic.List[string]

# Absolute path of the run log. Null until Initialize-AuditLog is called;
# logging before that point goes to the console only.
$script:AuditLogPath = $null

# Patterns that mask secret-shaped text even if the literal was never
# registered - for example a token quoted inside an exception message.
$script:AuditRedactionPatterns = @(
    # Bearer / CwsAuth token values
    '(?i)(Bearer[=\s]+)([A-Za-z0-9\-\._~\+/]{16,}=*)',
    # Form-encoded client_secret
    '(?i)(client_secret=)([^&\s]+)',
    # JSON access_token
    '(?i)("access_token"\s*:\s*")([^"]+)(")',
    # Basic auth header
    '(?i)(Basic\s+)([A-Za-z0-9\+/]{16,}=*)'
)

function Register-AuditSecret {
    <#
    .SYNOPSIS
        Registers a literal string that must never appear in log output.
    .DESCRIPTION
        Called as soon as a secret enters the process, so that any later
        accidental logging of it is masked. Values shorter than 4 characters
        are ignored, because masking them would corrupt ordinary text.
    #>
    [CmdletBinding()]
    param([string] $Secret)

    if ([string]::IsNullOrWhiteSpace($Secret)) { return }
    if ($Secret.Length -lt 4) { return }
    if (-not $script:AuditSecrets.Contains($Secret)) {
        $script:AuditSecrets.Add($Secret)
    }
}

function Get-RedactedText {
    <#
    .SYNOPSIS
        Masks registered secrets and secret-shaped substrings in text.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([AllowNull()][string] $Text)

    if ([string]::IsNullOrEmpty($Text)) { return '' }

    $result = $Text

    # Registered literals first - the exact values we know are secret.
    foreach ($secret in $script:AuditSecrets) {
        $result = $result.Replace($secret, '***REDACTED***')
    }

    # Then shape-based patterns, preserving the identifying prefix so the log
    # still shows what kind of value was masked.
    foreach ($pattern in $script:AuditRedactionPatterns) {
        $result = [regex]::Replace($result, $pattern, {
            param($m)
            if ($m.Groups.Count -ge 4 -and $m.Groups[3].Success) {
                $m.Groups[1].Value + '***REDACTED***' + $m.Groups[3].Value
            } else {
                $m.Groups[1].Value + '***REDACTED***'
            }
        })
    }

    return $result
}

function Initialize-AuditLog {
    <#
    .SYNOPSIS
        Opens the run log file, creating its directory if necessary.
    .DESCRIPTION
        Attempts to initialize the audit log file. On failure, leaves
        $script:AuditLogPath as $null so that Write-AuditLog degrades
        to console-only logging rather than failing repeatedly.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string] $Path)

    try {
        $dir = Split-Path -Parent $Path
        if ($dir -and -not (Test-Path $dir)) {
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
        }

        $header = @(
            '============================================================',
            ' Citrix Usage Report - run log',
            " Started (UTC): $([DateTime]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ssZ'))",
            " Host          : $($env:COMPUTERNAME)",
            " User          : $($env:USERDOMAIN)\$($env:USERNAME)",
            " PowerShell    : $($PSVersionTable.PSVersion)",
            '============================================================'
        )
        Set-Content -Path $Path -Value $header -Encoding UTF8 -ErrorAction Stop

        # Only set the path on successful write
        $script:AuditLogPath = $Path
    } catch {
        Write-Host "WARNING: Could not initialize run log at '$Path': $($_.Exception.Message)" -ForegroundColor Yellow
        # Leave $script:AuditLogPath as $null to degrade to console-only logging
    }
}

function Reset-AuditLogForTesting {
    <#
    .SYNOPSIS
        Clears logging state. Used by the test suite only.
    #>
    [CmdletBinding()]
    param()
    $script:AuditLogPath = $null
    $script:AuditSecrets.Clear()
}

function Write-AuditLog {
    <#
    .SYNOPSIS
        Writes a redacted, timestamped message to the console and the log file.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string] $Message,

        [ValidateSet('Debug', 'Info', 'Success', 'Warn', 'Error')]
        [string] $Level = 'Info'
    )

    $safe = Get-RedactedText -Text $Message
    $stamp = [DateTime]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ssZ')
    $line = '{0}  [{1,-7}] {2}' -f $stamp, $Level.ToUpperInvariant(), $safe

    switch ($Level) {
        'Debug'   { Write-Verbose $safe }
        'Info'    { Write-Host $safe }
        'Success' { Write-Host $safe -ForegroundColor Green }
        'Warn'    { Write-Host $safe -ForegroundColor Yellow }
        'Error'   { Write-Host $safe -ForegroundColor Red }
    }

    if ($script:AuditLogPath) {
        # Logging must never be the reason a run fails, so a write failure is
        # swallowed rather than propagated.
        try {
            Add-Content -Path $script:AuditLogPath -Value $line -Encoding UTF8 -ErrorAction Stop
        } catch {
            Write-Verbose "Could not write to log file: $($_.Exception.Message)"
        }
    }
}

# endregion 10-Logging.ps1

# ----------------------------------------------------------------------------
# region 20-Config.ps1
# ----------------------------------------------------------------------------
# ============================================================================
#  Configuration
#
#  Builds and validates the run configuration, and owns the mapping from a
#  chosen environment to concrete API endpoints. Nothing else in the script
#  hard-codes a Citrix URL.
# ============================================================================

# Citrix's own guidance on securing the Monitor Service with a TLS
# certificate. Linked from the unencrypted-HTTP notice in the console, the
# run log, and the HTML report -- one place to keep it in sync.
$script:MonitorTlsGuidanceUrl = 'https://developer-docs.citrix.com/en-us/monitor-service-odata-api/on-prem-odata.html'

# Endpoint table. The cloud region is always an explicit user choice: inferring
# it risks sending a Government tenant's credentials to a commercial endpoint.
$script:CitrixEndpoints = @{
    CloudCommercial = @{
        ODataBase = 'https://api.cloud.com/monitorodata'
        TokenBase = 'https://api.cloud.com'
        Label     = 'Citrix Cloud (Commercial - US / EU / AP-S)'
    }
    CloudJapan = @{
        ODataBase = 'https://api.citrixcloud.jp/monitorodata'
        TokenBase = 'https://api.citrixcloud.jp'
        Label     = 'Citrix Cloud (Japan)'
    }
    CloudGovernment = @{
        ODataBase = 'https://api.cloud.us/monitorodata'
        TokenBase = 'https://api.cloud.us'
        Label     = 'Citrix Cloud Government (US Gov)'
    }
}

function Resolve-CitrixEndpoint {
    <#
    .SYNOPSIS
        Maps an environment selection to concrete API endpoints.
    .DESCRIPTION
        On-premises, the Monitor Service can be published over HTTP or
        HTTPS -- that is a setting on the Delivery Controller's IIS site,
        not something this tool can infer. -Protocol is the caller's explicit
        choice, defaulting to Https. An explicit scheme typed into the
        hostname itself (a customer who pasted a browser URL, or one who
        already worked out their site is HTTP-only) always wins over
        -Protocol: it is more specific information than a default or a
        control the caller may not have touched.
    .OUTPUTS
        PSCustomObject with ODataBase, TokenBase, IsCloud, IsHttp, EnvironmentLabel.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)]
        [ValidateSet('OnPremises', 'CloudCommercial', 'CloudJapan', 'CloudGovernment')]
        [string] $Environment,

        [AllowEmptyString()]
        [string] $DeliveryController,

        [ValidateSet('Https', 'Http')]
        [string] $Protocol = 'Https'
    )

    if ($Environment -eq 'OnPremises') {
        if ([string]::IsNullOrWhiteSpace($DeliveryController)) {
            throw 'A DeliveryController hostname is required for on-premises environments.'
        }

        # Users routinely paste a full browser URL such as
        # https://ddc01.contoso.com/Director/. Reduce it to a bare host.
        # NOTE: not $host - that is a read-only PowerShell automatic variable.
        $hostName = $DeliveryController.Trim()

        # Schemes are case-insensitive per RFC 3986. An explicit scheme here
        # overrides -Protocol -- see the function's .DESCRIPTION.
        $scheme = $Protocol.ToLowerInvariant()
        $schemeMatch = [regex]::Match($hostName, '^\s*(https?)://',
            [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        if ($schemeMatch.Success) {
            $scheme = $schemeMatch.Groups[1].Value.ToLowerInvariant()
        }
        $hostName = [regex]::Replace($hostName, '^\s*https?://', '',
            [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)

        # A UNC path is not a hostname; fail clearly rather than build a malformed URL.
        if ($hostName -match '^\\\\') {
            throw "'$DeliveryController' looks like a UNC path, not a Citrix Delivery Controller hostname. Supply a hostname or FQDN, for example ddc01.contoso.com."
        }

        # Drop any userinfo component. A pasted credential must never reach the run log
        # or the HTML report by way of EnvironmentLabel.
        if ($hostName.Contains('@')) {
            $hostName = $hostName.Substring($hostName.LastIndexOf('@') + 1)
        }

        # Cut at the first path, query, or fragment delimiter.
        $hostName = ($hostName -split '[/?#]')[0]
        $hostName = $hostName.TrimEnd('.')

        if ([string]::IsNullOrWhiteSpace($hostName)) {
            throw 'A DeliveryController hostname is required for on-premises environments.'
        }

        return [pscustomobject]@{
            ODataBase        = '{0}://{1}/Citrix/Monitor/OData/v4/Data' -f $scheme, $hostName
            TokenBase        = $null
            IsCloud          = $false
            IsHttp           = ($scheme -eq 'http')
            EnvironmentLabel = "On-premises CVAD ($hostName)"
        }
    }

    $entry = $script:CitrixEndpoints[$Environment]
    return [pscustomobject]@{
        ODataBase        = $entry.ODataBase
        TokenBase        = $entry.TokenBase
        IsCloud          = $true
        IsHttp           = $false
        EnvironmentLabel = $entry.Label
    }
}

function New-AuditConfig {
    <#
    .SYNOPSIS
        Builds the run configuration object used by every later stage.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [ValidateSet('OnPremises', 'CloudCommercial', 'CloudJapan', 'CloudGovernment')]
        [string] $Environment = 'OnPremises',

        [AllowEmptyString()][string] $DeliveryController,
        [ValidateSet('Https', 'Http')][string] $Protocol = 'Https',
        [AllowEmptyString()][string] $CustomerId,
        [AllowEmptyString()][string] $ClientId,
        [System.Security.SecureString] $ClientSecret,
        [System.Management.Automation.PSCredential] $Credential,

        [int[]] $Days = @(30, 60, 90),
        [Parameter(Mandatory)][string] $OutputPath,

        [int] $BusinessHourStart = 8,
        [int] $BusinessHourEnd = 18,
        [string] $DisplayTimeZoneId = ([System.TimeZoneInfo]::Local.Id),

        [switch] $IncludeDeliveryGroups,
        [switch] $IncludeApplications,
        [switch] $IncludeClientDevices,
        [switch] $IncludeTrend,
        [switch] $Anonymize,
        [switch] $ExportRawData,
        [switch] $DemoData
    )

    $windows = @($Days | Sort-Object -Unique)
    if (-not $windows) { $windows = @(30, 60, 90) }

    # A single run start is captured once and reused everywhere, so that every
    # window and every "still running" session resolves against the same
    # instant. Recomputing "now" per call would make results subtly inconsistent.
    $nowUtc = [DateTime]::UtcNow

    $endpoint = if ($DemoData) {
        [pscustomobject]@{
            ODataBase = 'demo://synthetic'
            TokenBase = $null
            IsCloud   = $false
            IsHttp    = $false
            EnvironmentLabel = 'Synthetic demo data (no Citrix environment contacted)'
        }
    } else {
        Resolve-CitrixEndpoint -Environment $Environment -DeliveryController $DeliveryController -Protocol $Protocol
    }

    [pscustomobject]@{
        Environment           = $Environment
        EnvironmentLabel      = $endpoint.EnvironmentLabel
        ODataBase             = $endpoint.ODataBase
        TokenBase             = $endpoint.TokenBase
        IsCloud               = $endpoint.IsCloud
        # Whether the resolved Monitor endpoint is unencrypted HTTP -- either
        # because -Protocol Http was chosen, or because the hostname carried
        # an explicit http:// scheme. Read by 99-Main.ps1 to log the warning
        # loudly, and by the HTML report to carry it into the document a
        # customer forwards on. Always false for cloud and demo endpoints.
        IsHttp                = [bool] $endpoint.IsHttp

        DeliveryController    = $DeliveryController
        CustomerId            = $CustomerId
        ClientId              = $ClientId
        ClientSecret          = $ClientSecret
        Credential            = $Credential

        Days                  = $windows
        RunStartUtc           = $nowUtc
        WindowEndUtc          = $nowUtc
        WidestWindowStartUtc  = $nowUtc.AddDays(-1 * ($windows | Measure-Object -Maximum).Maximum)

        OutputPath            = $OutputPath
        BusinessHourStart     = $BusinessHourStart
        BusinessHourEnd       = $BusinessHourEnd
        DisplayTimeZoneId     = $DisplayTimeZoneId

        IncludeDeliveryGroups = [bool] $IncludeDeliveryGroups
        IncludeApplications   = [bool] $IncludeApplications
        IncludeClientDevices  = [bool] $IncludeClientDevices
        IncludeTrend          = [bool] $IncludeTrend
        Anonymize             = [bool] $Anonymize
        ExportRawData         = [bool] $ExportRawData
        DemoData              = [bool] $DemoData
    }
}

function Test-AuditConfig {
    <#
    .SYNOPSIS
        Validates a configuration, throwing a message aimed at the person
        running the script rather than at a developer.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][pscustomobject] $Config)

    if ($Config.DemoData) { return }

    if ($Config.IsCloud) {
        if ([string]::IsNullOrWhiteSpace($Config.CustomerId)) {
            throw 'CustomerId is required for Citrix Cloud environments. Find it in the Citrix Cloud console under Identity and Access Management > API Access.'
        }
        if ([string]::IsNullOrWhiteSpace($Config.ClientId)) {
            throw 'ClientId is required for Citrix Cloud environments. Create a service principal under Identity and Access Management > API Access > Service principals in the Citrix Cloud console, and use its Service Principal ID.'
        }
        if (-not $Config.ClientSecret) {
            throw 'ClientSecret is required for Citrix Cloud environments. It is shown once when the service principal is created under Identity and Access Management > API Access > Service principals in the Citrix Cloud console -- if it has been lost, generate a new one.'
        }
    } else {
        if ([string]::IsNullOrWhiteSpace($Config.DeliveryController)) {
            throw 'DeliveryController is required for on-premises environments. Supply the hostname of a Citrix Delivery Controller -- the machine running the Citrix Monitor Service. On smaller deployments Director may be installed on the same machine, but a standalone Director web server will not work.'
        }
    }

    if (-not $Config.OutputPath) {
        throw 'OutputPath is required.'
    }

    if ($Config.BusinessHourStart -ge $Config.BusinessHourEnd) {
        throw "Business hours are invalid: start ($($Config.BusinessHourStart)) must be earlier than end ($($Config.BusinessHourEnd))."
    }
}

# endregion 20-Config.ps1

# ----------------------------------------------------------------------------
# region 30-Auth.ps1
# ----------------------------------------------------------------------------
# ============================================================================
#  Authentication
#
#  Produces an "auth context" that the OData client uses without knowing
#  whether it is talking to an on-premises Delivery Controller over Negotiate or to
#  Citrix Cloud with a bearer token. The context exposes GetHeaders() and
#  Refresh(), and nothing else.
# ============================================================================

# Cloud bearer tokens are valid for one hour. A large 90-day pull can run
# longer than that, so the token is refreshed early rather than at expiry.
$script:TokenRefreshMarginSeconds = 600

function ConvertFrom-SecureStringPlain {
    <#
    .SYNOPSIS
        Converts a SecureString to plain text for the duration of one HTTP call.
    .DESCRIPTION
        The unmanaged buffer is always freed, including on failure. The plain
        value is registered for log redaction by the caller.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Mandatory)][System.Security.SecureString] $Secure)

    $ptr = [System.IntPtr]::Zero
    try {
        $ptr = [System.Runtime.InteropServices.Marshal]::SecureStringToGlobalAllocUnicode($Secure)
        return [System.Runtime.InteropServices.Marshal]::PtrToStringUni($ptr)
    } finally {
        if ($ptr -ne [System.IntPtr]::Zero) {
            [System.Runtime.InteropServices.Marshal]::ZeroFreeGlobalAllocUnicode($ptr)
        }
    }
}

function Request-CitrixCloudToken {
    <#
    .SYNOPSIS
        Exchanges a Citrix Cloud API client ID and secret for a bearer token.
    .OUTPUTS
        PSCustomObject with AccessToken and ExpiresUtc.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][pscustomobject] $Config)

    $tokenUrl = '{0}/cctrustoauth2/{1}/tokens/clients' -f $Config.TokenBase, [uri]::EscapeDataString($Config.CustomerId)
    $secret = ConvertFrom-SecureStringPlain -Secure $Config.ClientSecret

    # Register before use, so that even an exception quoting the request body
    # cannot leak the secret into the log. The body URL-encodes the secret,
    # which is a different string whenever it contains a reserved character
    # (+ / = & %, routine in generated API secrets), so the encoded form must
    # be registered too - otherwise it can reach a log line verbatim with no
    # client_secret= prefix in front of it for the shape-based pattern to catch.
    Register-AuditSecret -Secret $secret
    Register-AuditSecret -Secret ([uri]::EscapeDataString($secret))

    $body = 'grant_type=client_credentials&client_id={0}&client_secret={1}' -f `
        [uri]::EscapeDataString($Config.ClientId), [uri]::EscapeDataString($secret)

    Write-AuditLog -Level Debug -Message "Requesting Citrix Cloud token from $tokenUrl"

    try {
        $response = Invoke-RestMethod -Uri $tokenUrl -Method Post -Body $body `
            -ContentType 'application/x-www-form-urlencoded' `
            -Headers @{ Accept = 'application/json' } `
            -UseBasicParsing -TimeoutSec 60 -ErrorAction Stop
    } catch {
        $status = Get-ODataStatusCode -ErrorRecord $_
        if ($status -eq 401 -or $status -eq 400) {
            throw "Citrix Cloud rejected the credentials (HTTP $status). Check the customer ID, service principal ID and service principal secret, and confirm the service principal has not been revoked -- or its secret rotated, which is a common cause of a run that previously worked suddenly failing. Create or review service principals in the Citrix Cloud console under Identity and Access Management > API Access > Service principals."
        }
        throw "Could not obtain a Citrix Cloud token from $tokenUrl : $($_.Exception.Message)"
    } finally {
        $secret = $null
    }

    if (-not $response.access_token) {
        throw "The token endpoint $tokenUrl returned no access_token."
    }

    Register-AuditSecret -Secret $response.access_token

    $lifetime = 3600
    if ($response.expires_in) { $lifetime = [int] $response.expires_in }

    # Subtract the margin so the token is replaced before anything using it
    # can fail mid-pull.
    $expires = [datetime]::UtcNow.AddSeconds($lifetime - $script:TokenRefreshMarginSeconds)

    Write-AuditLog -Level Debug -Message "Token acquired; scheduled for refresh at $($expires.ToString('u'))."

    [pscustomobject]@{ AccessToken = $response.access_token; ExpiresUtc = $expires }
}

function New-CitrixAuthContext {
    <#
    .SYNOPSIS
        Builds the auth context used for every OData request.
    .OUTPUTS
        An object exposing GetHeaders() and Refresh().
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][pscustomobject] $Config)

    $ctx = [pscustomobject]@{
        ODataBase             = $Config.ODataBase
        IsCloud               = [bool] $Config.IsCloud
        CustomerId            = $Config.CustomerId
        Config                = $Config
        AccessToken           = $null
        TokenExpiresUtc       = [datetime]::MaxValue
        UseDefaultCredentials = $false
        Credential            = $Config.Credential
    }

    if ($ctx.IsCloud) {
        $token = Request-CitrixCloudToken -Config $Config
        $ctx.AccessToken = $token.AccessToken
        $ctx.TokenExpiresUtc = $token.ExpiresUtc
    } else {
        # On-premises uses Negotiate/NTLM. With no explicit credential the
        # identity of the account running the script is used, which is the
        # normal case for an administrator running this on a Delivery Controller.
        $ctx.UseDefaultCredentials = -not $Config.Credential
    }

    $ctx | Add-Member -MemberType ScriptMethod -Name Refresh -Value {
        if (-not $this.IsCloud) { return }
        $token = Request-CitrixCloudToken -Config $this.Config
        $this.AccessToken = $token.AccessToken
        $this.TokenExpiresUtc = $token.ExpiresUtc
    }

    $ctx | Add-Member -MemberType ScriptMethod -Name GetHeaders -Value {
        $headers = @{
            'Accept'       = 'application/json;odata.metadata=minimal'
            'User-Agent'   = 'CitrixUsageReport/1.0'
        }

        if ($this.IsCloud) {
            # Refresh proactively so a long pull never fails on an expired token.
            if ([datetime]::UtcNow -ge $this.TokenExpiresUtc) {
                Write-AuditLog -Level Debug -Message 'Bearer token near expiry; refreshing.'
                $this.Refresh()
            }
            $headers['Authorization'] = "CwsAuth Bearer=$($this.AccessToken)"
            $headers['Citrix-CustomerId'] = $this.CustomerId
        }

        return $headers
    }

    return $ctx
}

function Get-AuthRequestParameters {
    <#
    .SYNOPSIS
        Extra Invoke-RestMethod parameters that cannot be expressed as headers.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)] $AuthContext)

    $params = @{}
    if (-not $AuthContext.IsCloud) {
        if ($AuthContext.Credential) { $params['Credential'] = $AuthContext.Credential }
        else { $params['UseDefaultCredentials'] = $true }
    }
    return $params
}

# endregion 30-Auth.ps1

# ----------------------------------------------------------------------------
# region 40-ODataClient.ps1
# ----------------------------------------------------------------------------
# ============================================================================
#  OData client
#
#  Generic transport for the Citrix Monitor Service OData v4 API. It knows
#  about HTTP, paging and retry; it knows nothing about Citrix entities. The
#  AuthContext it receives supplies headers and can refresh itself, which keeps
#  the on-premises and cloud differences out of this file entirely.
# ============================================================================

# The Monitor API returns at most 1000 records per page, which is why
# Invoke-CitrixODataQuery must follow @odata.nextLink rather than assume a
# single page holds everything.

# Status codes worth retrying. Anything else is a permanent failure and
# retrying it only delays the error the user needs to see.
$script:ODataRetryableStatus = @(408, 429, 500, 502, 503, 504)

# Records every fetch that came back knowingly incomplete. Truncation used to
# be written to audit.log and nowhere else, so an under-counted device or
# application figure reached the customer's report looking exactly like a
# complete one -- the same silent-undercount failure the retention banner
# exists to prevent, applied to a different entity. The log is script-scoped
# because Invoke-CitrixODataQuery has no channel back to its caller for a
# second, out-of-band result; Get-CitrixDataset resets it before a run and
# reads it afterwards onto the dataset.
$script:ODataFetchWarnings = New-Object System.Collections.Generic.List[object]

function Reset-ODataFetchWarnings {
    <#
    .SYNOPSIS
        Clears the incomplete-fetch log at the start of a run.
    #>
    [CmdletBinding()]
    param()
    $script:ODataFetchWarnings = New-Object System.Collections.Generic.List[object]
}

function Add-ODataFetchWarning {
    <#
    .SYNOPSIS
        Records that one entity's data is knowingly incomplete or unbounded.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string] $Entity,
        [Parameter(Mandatory)][string] $Reason,
        [Parameter(Mandatory)][string] $Message,
        [int] $RecordCount = 0
    )

    $script:ODataFetchWarnings.Add([pscustomobject]@{
        Entity      = $Entity
        Reason      = $Reason
        RecordCount = $RecordCount
        Message     = $Message
    })
}

function Get-ODataFetchWarnings {
    <#
    .SYNOPSIS
        The incomplete-fetch log for the current run.
    .OUTPUTS
        A plain array. Callers wrap the call in @(), per the codebase
        convention -- see the note on Invoke-CitrixODataQuery's return.
    #>
    [CmdletBinding()]
    [OutputType([object[]])]
    param()
    return $script:ODataFetchWarnings.ToArray()
}

function New-ODataHttpError {
    <#
    .SYNOPSIS
        Builds an exception carrying an HTTP status code.
    .DESCRIPTION
        Used by the retry logic to classify failures, and by the test suite to
        simulate server responses without a network.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][int] $StatusCode, [string] $Message)

    if (-not $Message) { $Message = "The remote server returned HTTP $StatusCode." }
    $ex = [System.Exception]::new($Message)
    $ex.Data['StatusCode'] = $StatusCode
    return $ex
}

function Get-ODataStatusCode {
    <#
    .SYNOPSIS
        Extracts an HTTP status code from whatever error shape PowerShell 5.1
        produced.
    .DESCRIPTION
        Windows PowerShell surfaces web failures as WebException wrapped in
        varying ways depending on the call path, so several shapes are probed.
        Returns 0 when no status code is available, meaning "not an HTTP error".
    #>
    [CmdletBinding()]
    [OutputType([int])]
    param([Parameter(Mandatory)] $ErrorRecord)

    $ex = if ($ErrorRecord -is [System.Management.Automation.ErrorRecord]) {
        $ErrorRecord.Exception
    } else { $ErrorRecord }

    # A real InnerException chain is 2-4 deep. A cycle should never occur, but
    # walking one unguarded would hang inside error handling - the worst place
    # for a hang - so cap the depth defensively.
    $depth = 0
    $maxDepth = 32
    while ($ex -and $depth -lt $maxDepth) {
        if ($ex.Data -and $ex.Data['StatusCode']) { return [int] $ex.Data['StatusCode'] }
        if ($ex -is [System.Net.WebException] -and $ex.Response) {
            return [int] ([System.Net.HttpWebResponse] $ex.Response).StatusCode
        }
        if ($ex.PSObject.Properties['Response'] -and $ex.Response -and
            $ex.Response.PSObject.Properties['StatusCode']) {
            return [int] $ex.Response.StatusCode
        }
        $ex = $ex.InnerException
        $depth++
    }
    return 0
}

function Test-CitrixTlsTrustFailure {
    <#
    .SYNOPSIS
        True when a failure is a TLS certificate trust failure, as opposed to
        any other connection or HTTP error.
    .DESCRIPTION
        Windows PowerShell 5.1 runs on .NET Framework, where a certificate
        this machine does not trust surfaces as a WebException whose Status
        is WebExceptionStatus.TrustFailure, usually wrapping an
        AuthenticationException. That is checked first because it is a real
        typed signal, not a string match; the message-text check beneath it
        is a fallback for any wrapping shape that loses the typed exception
        on the way up.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param([Parameter(Mandatory)] $ErrorRecord)

    $ex = if ($ErrorRecord -is [System.Management.Automation.ErrorRecord]) {
        $ErrorRecord.Exception
    } else { $ErrorRecord }

    $depth = 0
    $maxDepth = 32
    while ($ex -and $depth -lt $maxDepth) {
        if ($ex -is [System.Net.WebException] -and
            $ex.Status -eq [System.Net.WebExceptionStatus]::TrustFailure) {
            return $true
        }
        if ($ex -is [System.Security.Authentication.AuthenticationException]) {
            return $true
        }
        if ($ex.Message -match '(?i)trust relationship|remote certificate is invalid|certificate.*not trusted') {
            return $true
        }
        $ex = $ex.InnerException
        $depth++
    }
    return $false
}

function New-CitrixTlsTrustMessage {
    <#
    .SYNOPSIS
        Explains a TLS certificate trust failure in terms a customer's
        administrator can act on.
    .DESCRIPTION
        The raw .NET exception -- "The underlying connection was closed:
        Could not establish trust relationship for the SSL/TLS secure
        channel." -- names the symptom and nothing else. The design spec
        promised a distinct diagnostic here; this is it.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Mandatory)][string] $Url)

    $targetHost = try { ([uri] $Url).Host } catch { $Url }

    "TLS certificate trust failure connecting to $targetHost. This machine does not trust the certificate presented by $targetHost. The usual cause is a certificate issued by the organisation's own internal Certificate Authority, whose root certificate is not present in this machine's trust store -- this is common when this tool is run from a machine outside the Citrix estate. Options: run this tool from a domain-joined machine inside the estate (which usually already trusts the internal CA), or install the issuing CA's root certificate on this machine. This tool does not offer a way to bypass certificate validation."
}

function Test-CitrixDirectorHost {
    <#
    .SYNOPSIS
        Cheap, bounded probe: does this host serve the Citrix Director web
        application?
    .DESCRIPTION
        Called only after the Monitor OData path has already 404'd, to turn a
        bare 404 into an actionable diagnosis. Citrix Director and the
        Monitor Service are different products: Director is a separate IIS
        web application that queries the Monitor Service, which runs on the
        Delivery Controller. On a small deployment they are commonly the same
        box, which is exactly why the distinction is otherwise invisible.

        One request, a short timeout, and every failure here is swallowed as
        "not Director" -- this probe must never be the thing that turns a
        clear 404 into a hang or a confusing secondary error.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param([Parameter(Mandatory)][string] $ODataUrl)

    try {
        $uri = [uri] $ODataUrl
        $directorUrl = '{0}://{1}/Director/' -f $uri.Scheme, $uri.Authority
        $response = Invoke-WebRequest -Uri $directorUrl -Method Get -UseBasicParsing `
            -TimeoutSec 10 -ErrorAction Stop
        return ([int] $response.StatusCode -eq 200)
    } catch {
        return $false
    }
}

function New-CitrixNotFoundMessage {
    <#
    .SYNOPSIS
        Builds the customer-facing message for a 404 on the Monitor OData
        path, probing for a Director-not-Monitor-Service misconfiguration
        before falling back to the generic message.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Mandatory)][string] $Url)

    if (Test-CitrixDirectorHost -ODataUrl $Url) {
        return "The Monitor OData endpoint was not found (HTTP 404) at $Url. This host is running Citrix Director but is not hosting the Monitor Service. The Monitor OData API runs on a Delivery Controller. Point this tool at a Delivery Controller instead -- on smaller deployments this may be the same machine, but on a split deployment it is a different one."
    }

    $protocolHint = if ($Url -match '^(?i)https://') {
        ' If the Monitor Service on this Delivery Controller is published over HTTP rather than HTTPS, re-run with -Protocol Http (or select Http in the dialog).'
    } else { '' }

    return "The Monitor OData endpoint was not found (HTTP 404) at $Url. Confirm the server named is a Delivery Controller and that the Monitor Service is running.$protocolHint"
}

function ConvertTo-ODataDateLiteral {
    <#
    .SYNOPSIS
        Formats a datetime as the ISO-8601 literal the Monitor API accepts.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Mandatory)][datetime] $Utc)

    # A local or unspecified time would shift every window boundary by the
    # machine's offset, so convert defensively.
    if ($Utc.Kind -ne [System.DateTimeKind]::Utc) { $Utc = $Utc.ToUniversalTime() }
    return $Utc.ToString('yyyy-MM-ddTHH:mm:ssZ')
}

function Build-ODataUrl {
    <#
    .SYNOPSIS
        Composes an OData query URL from its parts.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)][string] $Base,
        [Parameter(Mandatory)][string] $Entity,
        [string[]] $Select,
        [string] $Filter,
        [string] $OrderBy,
        [int] $Top
    )

    $url = "{0}/{1}" -f $Base.TrimEnd('/'), $Entity
    $parts = @()

    if ($Select)  { $parts += '$select=' + ($Select -join ',') }
    if ($Filter)  { $parts += '$filter=' + [uri]::EscapeDataString($Filter) }
    if ($OrderBy) { $parts += '$orderby=' + [uri]::EscapeDataString($OrderBy) }
    if ($Top -gt 0) { $parts += '$top=' + $Top }

    if ($parts.Count) { $url += '?' + ($parts -join '&') }
    return $url
}

function Invoke-CitrixODataQuery {
    <#
    .SYNOPSIS
        Runs an OData query, following paging and retrying transient failures.
    .OUTPUTS
        A plain array of the records from every page. Wrap the call in @()
        at the call site -- see the comment above the return statement.
    #>
    [CmdletBinding()]
    [OutputType([object[]])]
    param(
        [Parameter(Mandatory)] $AuthContext,
        [Parameter(Mandatory)][string] $Entity,
        [string[]] $Select,
        [string] $Filter,
        [string] $OrderBy,
        [int] $Top,
        [int] $MaxRecords = 0,
        [scriptblock] $ProgressAction,
        [int] $MaxAttempts = 5,
        [scriptblock] $CancellationCheck,

        # Defensive backstop against a server that never terminates paging.
        # A few thousand pages is far beyond any real Citrix site at 1000
        # records per page.
        [int] $MaxPages = 5000
    )

    $results = New-Object System.Collections.Generic.List[object]
    $url = Build-ODataUrl -Base $AuthContext.ODataBase -Entity $Entity `
        -Select $Select -Filter $Filter -OrderBy $OrderBy -Top $Top
    $page = 0
    $refreshedOn401 = $false

    # Tracks every page URL already requested. A server that returns a
    # non-advancing @odata.nextLink would otherwise loop forever - on a
    # customer's large pull that hang is indistinguishable from "still
    # working" until they kill the script.
    $seenUrls = New-Object 'System.Collections.Generic.HashSet[string]'

    Write-AuditLog -Level Debug -Message "OData GET $Entity"

    while ($url) {

        if ($CancellationCheck -and (& $CancellationCheck)) {
            Write-AuditLog -Level Warn -Message "Cancelled while fetching $Entity after $($results.Count) records."
            break
        }

        [void] $seenUrls.Add($url)

        $attempt = 0
        $response = $null

        while ($true) {
            $attempt++
            try {
                $extra = Get-AuthRequestParameters -AuthContext $AuthContext
                $response = Invoke-RestMethod -Uri $url -Method Get `
                    -Headers $AuthContext.GetHeaders() -UseBasicParsing `
                    -TimeoutSec 300 -ErrorAction Stop @extra
                break
            } catch {
                $status = Get-ODataStatusCode -ErrorRecord $_

                # A 401 usually means an expired cloud token. Refresh once and
                # retry; a second 401 is a real credential problem.
                if ($status -eq 401 -and -not $refreshedOn401) {
                    $refreshedOn401 = $true
                    Write-AuditLog -Level Debug -Message 'Received HTTP 401; refreshing credentials and retrying.'
                    $AuthContext.Refresh()
                    continue
                }

                if ($status -eq 401) {
                    # An interactive run that never asked for credentials
                    # authenticated as the signed-in user by default (see
                    # Get-AuthRequestParameters in 30-Auth.ps1, which is what
                    # AuthContext.UseDefaultCredentials reflects). A reader
                    # who was never asked for credentials has no reason to
                    # suspect that is what went wrong unless the message says
                    # so, and names the three places a different account can
                    # be supplied instead.
                    if (-not $AuthContext.IsCloud -and $AuthContext.UseDefaultCredentials) {
                        throw "Authentication failed (HTTP 401) against $($AuthContext.ODataBase). This run authenticated as the signed-in user, which is the default. Check that account may read Citrix Monitor data, or supply a different account -- in the dialog, at the console prompt, or with -Credential."
                    }
                    throw "Authentication failed (HTTP 401) against $($AuthContext.ODataBase). Check the credentials supplied and that the account may read Citrix Monitor data."
                }

                if ($status -eq 403) {
                    throw "Access denied (HTTP 403) against $($AuthContext.ODataBase). The account authenticated successfully but lacks permission to read Citrix Monitor data. On-premises this needs a Citrix Director or Monitor read role; in Citrix Cloud the API client needs a read scope."
                }

                if ($status -eq 404) {
                    throw (New-CitrixNotFoundMessage -Url $url)
                }

                # Checked ahead of the generic retryable/exhausted branch below
                # because it is never retryable (the same certificate will be
                # rejected every time) and, unhandled, surfaced as the raw
                # .NET exception text -- "The underlying connection was
                # closed: Could not establish trust relationship for the
                # SSL/TLS secure channel." -- which names the symptom and
                # gives the reader nothing to act on.
                if (Test-CitrixTlsTrustFailure -ErrorRecord $_) {
                    throw (New-CitrixTlsTrustMessage -Url $url)
                }

                $retryable = $script:ODataRetryableStatus -contains $status
                if (-not $retryable -or $attempt -ge $MaxAttempts) {
                    $detail = if ($status) { "HTTP $status" } else { $_.Exception.Message }
                    # Thrown as an exception carrying the status code rather
                    # than as a bare string, so a caller can distinguish "the
                    # server rejected my $filter" (400) from a real failure
                    # without string-matching the message. The message text is
                    # unchanged.
                    throw (New-ODataHttpError -StatusCode $status `
                        -Message "Query for $Entity failed after $attempt attempt(s): $detail")
                }

                # Exponential backoff, capped so a long outage does not stall
                # the run for minutes at a time.
                $delay = [math]::Min([math]::Pow(2, $attempt), 30)
                Write-AuditLog -Level Warn -Message "HTTP $status fetching $Entity; retrying in $delay s (attempt $attempt of $MaxAttempts)."
                Start-Sleep -Seconds $delay
            }
        }

        $page++
        $batch = @($response.value)
        foreach ($item in $batch) { [void] $results.Add($item) }

        if ($ProgressAction) { & $ProgressAction $Entity $results.Count }

        if ($MaxRecords -gt 0 -and $results.Count -ge $MaxRecords) {
            Write-AuditLog -Level Warn -Message "Reached the $MaxRecords record cap for $Entity; results are truncated."
            Add-ODataFetchWarning -Entity $Entity -Reason 'RecordCap' -RecordCount $results.Count `
                -Message "Only the first $($results.Count) $Entity record(s) were read before the record cap was reached, and pages do not arrive in time order, so the records that were dropped are an arbitrary subset."
            break
        }

        if ($page -ge $MaxPages) {
            Write-AuditLog -Level Warn -Message "Reached the maximum of $MaxPages page(s) fetching $Entity without the server signalling completion; stopping to avoid an unbounded run. Results may be incomplete."
            Add-ODataFetchWarning -Entity $Entity -Reason 'PageCap' -RecordCount $results.Count `
                -Message "The server was still returning $Entity pages after $MaxPages page(s) ($($results.Count) record(s)), so the fetch was stopped. Pages do not arrive in time order, so the records that were dropped are an arbitrary subset."
            break
        }

        # OData v4 signals more data with @odata.nextLink. Ignoring it would
        # silently truncate every large site to its first 1000 records.
        $next = $response.PSObject.Properties['@odata.nextLink']
        if (-not $next) {
            $url = $null
        } elseif ($seenUrls.Contains($next.Value)) {
            # A server returning a non-advancing nextLink would otherwise loop
            # forever. Truncation must never be silent, so this is a warning,
            # not a quiet stop.
            Write-AuditLog -Level Warn -Message "The server returned a repeating pagination link while fetching $Entity; stopping to avoid an infinite loop. Results may be incomplete."
            Add-ODataFetchWarning -Entity $Entity -Reason 'RepeatingPageLink' -RecordCount $results.Count `
                -Message "The server returned a repeating pagination link while fetching $Entity, so the fetch was stopped after $($results.Count) record(s) to avoid an infinite loop."
            $url = $null
        } else {
            $url = $next.Value
        }
    }

    Write-AuditLog -Level Debug -Message "$Entity : $($results.Count) record(s) over $page page(s)."
    # Plain return, not a unary-comma-wrapped one: `@(Invoke-CitrixODataQuery
    # ...)` at the call site (see 50-DataModel.ps1) is the codebase's
    # standard way to get a reliably correct count for zero, one, or many
    # records. A comma-wrapped return works for a *bare* call but nests when
    # the caller applies that same @() -- .Count would then report 1
    # regardless of how many records actually came back, which is worse
    # than the scalar-collapse bug the comma was added to prevent, and is
    # also mock-incompatible: tests (and Task 6's preflight) mock this
    # function's return as a plain array.
    return $results.ToArray()
}

# endregion 40-ODataClient.ps1

# ----------------------------------------------------------------------------
# region 50-DataModel.ps1
# ----------------------------------------------------------------------------
# ============================================================================
#  Data model
#
#  Fetches Citrix Monitor entities and projects them into slim, normalised
#  objects. Projection happens as each page arrives so that a large site does
#  not accumulate raw JSON in memory.
#
#  Every later stage consumes the dataset shape produced here, so the demo
#  generator emits exactly the same shape as the live fetchers.
# ============================================================================

function ConvertTo-UtcDateTime {
    <#
    .SYNOPSIS
        Parses an OData date string into a UTC DateTime, or $null.
    .DESCRIPTION
        The Monitor API returns timestamps in several ISO-8601 shapes depending
        on version and field. Round-trip parsing normalises all of them, and a
        null or empty value maps to $null rather than to DateTime.MinValue,
        because "no end date" means the session is still running.
    #>
    [CmdletBinding()]
    param([AllowNull()] $Value)

    if ($null -eq $Value -or $Value -eq '') { return $null }
    if ($Value -is [datetime]) {
        return $(if ($Value.Kind -eq [System.DateTimeKind]::Utc) { $Value } else { $Value.ToUniversalTime() })
    }

    $parsed = [datetime]::MinValue
    $styles = [System.Globalization.DateTimeStyles]::AdjustToUniversal -bor
              [System.Globalization.DateTimeStyles]::AssumeUniversal
    if ([datetime]::TryParse([string] $Value, [System.Globalization.CultureInfo]::InvariantCulture,
            $styles, [ref] $parsed)) {
        return [datetime]::SpecifyKind($parsed, [System.DateTimeKind]::Utc)
    }
    return $null
}

function Get-SessionOverlapFilter {
    <#
    .SYNOPSIS
        Builds the OData filter selecting sessions that overlap a window.
    .DESCRIPTION
        Overlap, not start date. A session that began before the window still
        consumes a licence inside it, and persistent desktops routinely run for
        weeks. Filtering on StartDate alone would undercount every site that
        uses static desktops.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)][datetime] $StartUtc,
        [Parameter(Mandatory)][datetime] $EndUtc
    )

    $s = ConvertTo-ODataDateLiteral -Utc $StartUtc
    $e = ConvertTo-ODataDateLiteral -Utc $EndUtc
    return "(EndDate eq null or EndDate gt $s) and StartDate lt $e"
}

function Invoke-AuditPreflight {
    <#
    .SYNOPSIS
        Verifies reachability and discovers how much history actually exists.
    .DESCRIPTION
        Citrix grooms raw session data on a schedule set by the site's licence
        edition: 90 days on Premium, 31 on Advanced, 7 otherwise. Any window
        wider than the real limit produces an under-count that is invisible in
        the output unless it is labelled here.

        Retention is measured from the oldest ENDED session, because grooming
        only ever removes ended sessions. A persistent desktop that has been
        connected for months is never groomed, and it wins a plain
        "oldest session by StartDate" query outright -- so deriving retention
        that way lets ONE long-running session claim months of history that
        does not exist, and the report then affirms full coverage over windows
        it cannot cover. That is the single worst output this tool can
        produce, so the oldest-start probe is kept only as reported context
        and is never allowed to set AvailableDays on its own.

        Two caveats are deliberately carried in the returned object rather
        than papered over:

        * The oldest surviving ended session gives a LOWER bound on retention,
          not retention itself. On a quiet site the oldest ended session can be
          recent simply because nothing ended earlier -- not because anything
          was groomed. The figure is still the honest answer to "how far back
          does the data we hold actually go", which is what the windows are
          analysed against, so the banner wording names both possible causes
          instead of asserting grooming.
        * A site with NO ended sessions in retention gives no grooming signal
          at all. Rather than report zero (which would flag every window on a
          site whose sessions all happen to be open) the oldest-start probe is
          used as a fallback and RetentionBasis records that retention could
          not be confirmed.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] $AuthContext,
        [Parameter(Mandatory)][pscustomobject] $Config
    )

    Write-AuditLog -Level Info -Message 'Checking connectivity and available history...'

    # @() defensively re-wraps each result: PowerShell collapses a single-
    # element (or empty) array to a scalar (or $null) as it crosses a
    # function's output stream unless the producer guards against it (see the
    # unary-comma comment in Invoke-CitrixODataQuery). -Top 1 means these calls
    # are exactly the case most likely to return one record, so relying on the
    # producer's guard alone - which a test double does not reproduce - would
    # make $rows.Count silently resolve to $null here.
    $oldestStart = $null
    $startRows = @(Invoke-CitrixODataQuery -AuthContext $AuthContext -Entity 'Sessions' `
        -Select @('SessionKey', 'StartDate') -OrderBy 'StartDate asc' -Top 1)
    if ($startRows -and $startRows.Count -gt 0) {
        $oldestStart = ConvertTo-UtcDateTime -Value $startRows[0].StartDate
    }

    # The retention signal. "EndDate ne null" restricts the probe to sessions
    # that are actually subject to grooming; ordering by EndDate finds the
    # oldest one that survived it.
    $oldestEnd = $null
    $endedRows = @(Invoke-CitrixODataQuery -AuthContext $AuthContext -Entity 'Sessions' `
        -Select @('SessionKey', 'StartDate', 'EndDate') -Filter 'EndDate ne null' `
        -OrderBy 'EndDate asc' -Top 1)
    if ($endedRows -and $endedRows.Count -gt 0) {
        $oldestEnd = ConvertTo-UtcDateTime -Value $endedRows[0].EndDate
    }

    $historyStart = $null
    $basis = 'NoSessions'
    if ($oldestEnd) {
        $historyStart = $oldestEnd
        $basis = 'OldestEndedSession'
    } elseif ($oldestStart) {
        $historyStart = $oldestStart
        $basis = 'OldestSessionStart'
    }

    # Defensive clamp: a site clock ahead of this machine's could put either
    # probe in the future, which would make AvailableDays negative and, worse,
    # push the analysis window start past "now".
    if ($historyStart -and $historyStart -gt $Config.RunStartUtc) {
        $historyStart = $Config.RunStartUtc
    }

    $availableDays = 0
    if ($historyStart) {
        $availableDays = [math]::Floor(($Config.RunStartUtc - $historyStart).TotalDays)
        if ($availableDays -lt 0) { $availableDays = 0 }
    }

    $truncated = @($Config.Days | Where-Object { $_ -gt $availableDays })

    if ($basis -eq 'OldestSessionStart') {
        Write-AuditLog -Level Warn -Message 'No completed (ended) sessions were returned, so how much history Citrix still retains could not be confirmed. Falling back to the oldest session start date, which may overstate retention if every older session has already been groomed.'
    } elseif ($basis -eq 'NoSessions') {
        Write-AuditLog -Level Warn -Message 'No sessions at all were returned by the history probe. Either the site has had no activity, or the account cannot read Monitor session data.'
    }

    if ($truncated.Count -gt 0) {
        Write-AuditLog -Level Warn -Message ("Only {0} day(s) of session history are available on this site. These windows cannot be fully covered: {1}. Their figures are lower bounds, not true counts." -f $availableDays, ($truncated -join ', '))
        Write-AuditLog -Level Warn -Message 'Citrix grooms raw session data at 90 days on Premium, 31 on Advanced, and 7 on other editions. Increasing retention requires a Premium licence and a Set-MonitorConfiguration change. (History can also be short simply because the site had no activity further back.)'
    } elseif ($basis -ne 'OldestEndedSession') {
        # Every *requested* window nominally fits inside $availableDays, but
        # $availableDays itself was never confirmed against Citrix's grooming
        # signal (no ended session was found to measure it from) -- so this is
        # NOT the "fully covered" affirmative claim below. Saying that here
        # would be the exact false-confidence sentence this preflight exists
        # to prevent: a site with zero ended sessions in retention (a handful
        # of always-connected persistent desktops on a 7-day, non-Premium
        # site is a real, unremarkable example) would otherwise get a report
        # with no banner and no caveat, stating coverage that was never
        # actually established.
        Write-AuditLog -Level Warn -Message "$availableDays day(s) of session history appear to be available, but this could not be confirmed: no completed (ended) session was found to measure retention from. Every requested window nominally fits inside that figure, but treat it -- and everything derived from it -- as a lower bound, not a confirmed true count."
    } else {
        Write-AuditLog -Level Success -Message "$availableDays day(s) of session history are available; every requested window is fully covered."
    }

    if ($oldestStart -and $oldestEnd -and $oldestStart -lt $historyStart) {
        Write-AuditLog -Level Info -Message ("The oldest session on this site started {0:yyyy-MM-dd} and is still running; retention is measured from the oldest ENDED session ({1:yyyy-MM-dd}) because open sessions are never groomed." -f $oldestStart, $oldestEnd)
    }

    [pscustomobject]@{
        Reachable            = $true
        OldestSessionUtc     = $oldestStart
        OldestEndedSessionUtc= $oldestEnd
        HistoryStartUtc      = $historyStart
        RetentionBasis       = $basis
        AvailableDays        = [int] $availableDays
        TruncatedWindows     = $truncated
        CheckedAtUtc         = $Config.RunStartUtc
    }
}

function Get-CitrixSessionRecords {
    <#
    .SYNOPSIS
        Fetches sessions overlapping the widest requested window.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] $AuthContext,
        [Parameter(Mandatory)][pscustomobject] $Config,
        [scriptblock] $ProgressAction,
        [scriptblock] $CancellationCheck
    )

    $filter = Get-SessionOverlapFilter -StartUtc $Config.WidestWindowStartUtc -EndUtc $Config.WindowEndUtc

    $raw = Invoke-CitrixODataQuery -AuthContext $AuthContext -Entity 'Sessions' `
        -Select @('SessionKey','UserId','MachineId','StartDate','EndDate','SessionType','IsAnonymous','ConnectionState','LifecycleState') `
        -Filter $filter -ProgressAction $ProgressAction -CancellationCheck $CancellationCheck

    foreach ($r in $raw) {
        [pscustomobject]@{
            SessionKey  = [string] $r.SessionKey
            UserId      = [string] $r.UserId
            MachineId   = [string] $r.MachineId
            StartUtc    = ConvertTo-UtcDateTime -Value $r.StartDate
            EndUtc      = ConvertTo-UtcDateTime -Value $r.EndDate
            SessionType = [int] ($(if ($null -ne $r.SessionType) { $r.SessionType } else { 0 }))
            IsAnonymous = [bool] $r.IsAnonymous
        }
    }
}

function Get-CitrixLookupRecords {
    <#
    .SYNOPSIS
        Fetches a small lookup entity and projects the named fields.
    .DESCRIPTION
        Users, machines, delivery groups and catalogs are all small, flat
        lookups fetched the same way. One function avoids four near-identical
        ones.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] $AuthContext,
        [Parameter(Mandatory)][string] $Entity,
        [Parameter(Mandatory)][string[]] $Fields,
        [string] $Filter,
        [scriptblock] $ProgressAction
    )

    $raw = Invoke-CitrixODataQuery -AuthContext $AuthContext -Entity $Entity `
        -Select $Fields -Filter $Filter -ProgressAction $ProgressAction

    foreach ($r in $raw) {
        $o = [ordered]@{}
        foreach ($f in $Fields) { $o[$f] = $r.$f }
        [pscustomobject] $o
    }
}

function Get-SessionLinkedWindowFilter {
    <#
    .SYNOPSIS
        Bounds a session-linked entity (Connections, ApplicationInstances) to
        the period the sessions themselves cover.
    .DESCRIPTION
        A connection, and an application launch, both begin at or after the
        start of the session they belong to and before the end of the
        reporting window. So the earliest START of any session actually
        fetched is a LOSSLESS lower bound for both entities: no row belonging
        to any fetched session can predate it. That matters -- a naive
        window-start bound would drop the connections and launches of
        persistent desktops that began before the window, which are exactly
        the sessions the overlap filter goes out of its way to keep.

        Without a bound these entities pull the site's ENTIRE retained
        history rather than the requested period, which on a large site is
        millions of rows and can trip the paging cap -- truncating the device
        and application figures by an arbitrary subset.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)][string] $DateField,
        [Parameter(Mandatory)][datetime] $StartUtc,
        [Parameter(Mandatory)][datetime] $EndUtc
    )

    $s = ConvertTo-ODataDateLiteral -Utc $StartUtc
    $e = ConvertTo-ODataDateLiteral -Utc $EndUtc
    return "$DateField ge $s and $DateField lt $e"
}

function Get-CitrixWindowedRecords {
    <#
    .SYNOPSIS
        Fetches a session-linked entity bounded to the reporting window,
        falling back to an unbounded fetch if the site rejects the filter.
    .DESCRIPTION
        The date field these entities carry is NOT StartDate/EndDate as on
        Sessions, and it has varied across Monitor Service versions. This
        project has never been run against a live Citrix site, so the field
        names used by the caller are the best available reading of the schema
        rather than something verified in production.

        A wrong field name in a filter is an HTTP 400, which would otherwise
        turn a working (if wasteful) fetch into a failed audit. So a 400 --
        and only a 400 -- degrades to the previous unfiltered behaviour, logs
        loudly, and records a fetch warning that reaches the REPORT, not just
        audit.log. That way the bound is taken wherever the site supports it,
        and a site that does not support it behaves exactly as it did before
        with the risk stated on the face of the output.
    #>
    [CmdletBinding()]
    [OutputType([object[]])]
    param(
        [Parameter(Mandatory)] $AuthContext,
        [Parameter(Mandatory)][string] $Entity,
        [Parameter(Mandatory)][string[]] $Fields,
        [string] $Filter,
        [scriptblock] $ProgressAction
    )

    if ([string]::IsNullOrWhiteSpace($Filter)) {
        return @(Get-CitrixLookupRecords -AuthContext $AuthContext -Entity $Entity `
            -Fields $Fields -ProgressAction $ProgressAction)
    }

    try {
        return @(Get-CitrixLookupRecords -AuthContext $AuthContext -Entity $Entity `
            -Fields $Fields -Filter $Filter -ProgressAction $ProgressAction)
    } catch {
        if ((Get-ODataStatusCode -ErrorRecord $_) -ne 400) { throw }

        Write-AuditLog -Level Warn -Message "This site rejected the date filter used to limit the $Entity fetch to the reporting window ($Filter). Falling back to fetching all retained $Entity records, which is slower and, on a very large site, can hit the paging cap."
        Add-ODataFetchWarning -Entity $Entity -Reason 'UnboundedFetch' `
            -Message "$Entity could not be limited to the reporting period because this site rejected the date filter, so every retained $Entity record was requested instead. The figures derived from it are still correct unless a paging cap warning also appears."

        return @(Get-CitrixLookupRecords -AuthContext $AuthContext -Entity $Entity `
            -Fields $Fields -ProgressAction $ProgressAction)
    }
}

function Get-CitrixDataset {
    <#
    .SYNOPSIS
        Fetches every entity the configured toggles require.
    .DESCRIPTION
        Sessions and the small lookups are always fetched. Connections and
        application instances are the expensive entities on a large site and
        are fetched only when their breakdown is switched on.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] $AuthContext,
        [Parameter(Mandatory)][pscustomobject] $Config,
        [scriptblock] $ProgressAction,
        [scriptblock] $CancellationCheck
    )

    # Cleared before anything is fetched so the warnings this run collects
    # cannot include a previous run's (the log is script-scoped; see
    # src/40-ODataClient.ps1).
    Reset-ODataFetchWarnings

    $preflight = Invoke-AuditPreflight -AuthContext $AuthContext -Config $Config

    Write-AuditLog -Level Info -Message 'Fetching sessions...'
    $sessions = @(Get-CitrixSessionRecords -AuthContext $AuthContext -Config $Config `
        -ProgressAction $ProgressAction -CancellationCheck $CancellationCheck)
    Write-AuditLog -Level Success -Message "Fetched $($sessions.Count) session record(s)."

    Write-AuditLog -Level Info -Message 'Fetching users, machines and delivery groups...'
    $users = @(Get-CitrixLookupRecords -AuthContext $AuthContext -Entity 'Users' `
        -Fields @('Id','UserName','FullName','Upn','Sid','Domain') -ProgressAction $ProgressAction)
    $machines = @(Get-CitrixLookupRecords -AuthContext $AuthContext -Entity 'Machines' `
        -Fields @('Id','Name','DesktopGroupId','CatalogId') -ProgressAction $ProgressAction)
    $groups = @(Get-CitrixLookupRecords -AuthContext $AuthContext -Entity 'DesktopGroups' `
        -Fields @('Id','Name','SessionSupport','DeliveryType') -ProgressAction $ProgressAction)
    $catalogs = @(Get-CitrixLookupRecords -AuthContext $AuthContext -Entity 'Catalogs' `
        -Fields @('Id','Name') -ProgressAction $ProgressAction)

    # The lower bound both toggled entities are filtered on. Sessions have
    # already been fetched at this point, so the earliest session start is
    # known and is a lossless bound (see Get-SessionLinkedWindowFilter).
    # With no sessions at all there is nothing to join to, so the widest
    # requested window is used instead.
    $linkedStartUtc = $Config.WidestWindowStartUtc
    if ($sessions.Count -gt 0) {
        $earliest = ($sessions | Where-Object { $_.StartUtc } |
            Measure-Object -Property StartUtc -Minimum).Minimum
        if ($earliest) { $linkedStartUtc = $earliest }
    }

    $connections = @()
    if ($Config.IncludeClientDevices) {
        Write-AuditLog -Level Info -Message 'Fetching connections for the client device breakdown...'
        # EstablishmentDate, not StartDate: the Connection entity does not
        # carry the Session entity's date fields. See Get-CitrixWindowedRecords
        # for what happens if this site does not know that field.
        $connections = @(Get-CitrixWindowedRecords -AuthContext $AuthContext -Entity 'Connections' `
            -Fields @('Id','SessionKey','ClientName','ClientAddress','ClientVersion','IsReconnect') `
            -Filter (Get-SessionLinkedWindowFilter -DateField 'EstablishmentDate' `
                -StartUtc $linkedStartUtc -EndUtc $Config.WindowEndUtc) `
            -ProgressAction $ProgressAction)
    }

    $appInstances = @(); $applications = @()
    if ($Config.IncludeApplications) {
        Write-AuditLog -Level Info -Message 'Fetching application instances for the application breakdown...'
        # No date fields: nothing in this project reads them (only SessionKey and
        # ApplicationId are ever consumed), and Get-CitrixLookupRecords is a raw
        # pass-through that never applies ConvertTo-UtcDateTime, unlike Sessions -
        # so StartDate/EndDate here would stay ISO-8601 strings on the live path.
        # New-DemoDataset must not emit them either, or the demo dataset would
        # diverge in shape from live data (real [datetime] vs. raw string) for a
        # field nothing uses. See the matching note in New-DemoDataset.
        # StartDate here is ApplicationInstance's own launch time, which the
        # entity does carry even though this select list does not request it
        # (see the note above): it is used only to bound the fetch, never
        # read. Filtering on it does not put a raw ISO-8601 string on the
        # projected object, so the live/demo shape parity the test suite pins
        # is unaffected.
        $appInstances = @(Get-CitrixWindowedRecords -AuthContext $AuthContext -Entity 'ApplicationInstances' `
            -Fields @('Id','SessionKey','ApplicationId') `
            -Filter (Get-SessionLinkedWindowFilter -DateField 'StartDate' `
                -StartUtc $linkedStartUtc -EndUtc $Config.WindowEndUtc) `
            -ProgressAction $ProgressAction)
        $applications = @(Get-CitrixLookupRecords -AuthContext $AuthContext -Entity 'Applications' `
            -Fields @('Id','Name','PublishedName') -ProgressAction $ProgressAction)
    }

    [pscustomobject]@{
        Sessions             = $sessions
        Users                = $users
        Machines             = $machines
        DesktopGroups        = $groups
        Catalogs             = $catalogs
        Connections          = $connections
        ApplicationInstances = $appInstances
        Applications         = $applications
        Preflight            = $preflight
        # Fetches that came back knowingly incomplete or unbounded. Carried on
        # the dataset so the REPORT can say so next to the affected figures,
        # rather than the customer having to read audit.log to discover that a
        # device count is an arbitrary subset.
        FetchWarnings        = @(Get-ODataFetchWarnings)
        IsDemo               = $false
    }
}

function New-DemoDataset {
    <#
    .SYNOPSIS
        Generates a synthetic Citrix site with realistic usage shape.
    .DESCRIPTION
        Lets the whole pipeline be developed, tested and demonstrated without a
        Citrix environment. The shape deliberately includes the cases that
        break naive analytics: sessions still running, sessions that began
        before the window, weekend troughs and lunchtime peaks.

        Deterministic for a given seed, so it can back regression tests.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][pscustomobject] $Config,
        [int] $Seed = 20260821,
        [int] $UserCount = 250
    )

    $rand = [System.Random]::new($Seed)

    $groupNames = @('Finance Desktops','Engineering Desktops','Call Centre Apps','Executive Desktops','Contractor Apps')
    $groups = @(); $catalogs = @(); $machines = @()

    for ($g = 0; $g -lt $groupNames.Count; $g++) {
        $gid = "dg-$g"
        $groups += [pscustomobject]@{
            Id = $gid; Name = $groupNames[$g]
            SessionSupport = $(if ($g -in 0,1,3) { 1 } else { 2 })
            DeliveryType = 1
        }
        $catalogs += [pscustomobject]@{ Id = "cat-$g"; Name = "$($groupNames[$g]) Catalog" }
        $machineCount = 4 + $rand.Next(0, 8)
        for ($m = 0; $m -lt $machineCount; $m++) {
            $machines += [pscustomobject]@{
                Id = "mc-$g-$m"; Name = "VDA-$g-$([string]::Format('{0:D2}', $m))"
                DesktopGroupId = $gid; CatalogId = "cat-$g"
            }
        }
    }

    $users = @()
    for ($u = 0; $u -lt $UserCount; $u++) {
        $users += [pscustomobject]@{
            Id = "usr-$u"
            UserName = "user{0:D4}" -f $u
            FullName = "Demo User $u"
            Upn = "user{0:D4}@demo.local" -f $u
            Sid = "S-1-5-21-0-0-0-{0}" -f (1000 + $u)
            Domain = 'DEMO'
        }
    }

    # A slice of the population stops appearing after the first third of the
    # window, so the daily trend and per-window unique-user counts vary
    # realistically instead of holding a flat population throughout.
    $churnCutoff = [math]::Floor($UserCount * 0.18)

    $sessions = New-Object System.Collections.Generic.List[object]
    $windowStart = $Config.WidestWindowStartUtc
    $windowEnd = $Config.WindowEndUtc
    $totalDays = [math]::Ceiling(($windowEnd - $windowStart).TotalDays)
    $key = 0

    for ($d = 0; $d -lt $totalDays; $d++) {
        $day = $windowStart.Date.AddDays($d)
        $isWeekend = $day.DayOfWeek -in @([DayOfWeek]::Saturday, [DayOfWeek]::Sunday)

        # Weekends run at roughly a sixth of weekday volume.
        $activeShare = if ($isWeekend) { 0.10 } else { 0.62 }
        $activeCount = [int] ($UserCount * $activeShare * (0.85 + $rand.NextDouble() * 0.3))

        for ($i = 0; $i -lt $activeCount; $i++) {
            $userIndex = $rand.Next(0, $UserCount)

            # Churned users disappear after the first third of the window.
            if ($userIndex -lt $churnCutoff -and $d -gt ($totalDays / 3)) { continue }

            # Log-on times cluster on a morning peak with a lunchtime second
            # wave, which is what makes the concurrency curve look real.
            $hour = if ($rand.NextDouble() -lt 0.65) { 7 + $rand.Next(0, 3) } else { 11 + $rand.Next(0, 7) }
            $start = $day.AddHours($hour).AddMinutes($rand.Next(0, 60))
            $start = [datetime]::SpecifyKind($start, [System.DateTimeKind]::Utc)
            if ($start -ge $windowEnd) { continue }

            $durationHours = 1 + $rand.NextDouble() * 8
            $end = $start.AddHours($durationHours)

            # Persistent desktops that never log off, and sessions still open
            # at the moment of the run.
            $stillRunning = ($end -ge $windowEnd) -or ($rand.NextDouble() -lt 0.004)
            if ($stillRunning) { $end = $null }

            $machine = $machines[$rand.Next(0, $machines.Count)]
            $group = $groups | Where-Object { $_.Id -eq $machine.DesktopGroupId } | Select-Object -First 1

            $sessions.Add([pscustomobject]@{
                SessionKey  = "sess-$key"
                UserId      = "usr-$userIndex"
                MachineId   = $machine.Id
                StartUtc    = $start
                EndUtc      = $end
                SessionType = $(if ($group.SessionSupport -eq 1) { 0 } else { 1 })
                IsAnonymous = ($rand.NextDouble() -lt 0.005)
            })
            $key++
        }
    }

    # A handful of long-lived desktops that began before the widest window, so
    # the overlap filter and the window clamping are both exercised.
    for ($i = 0; $i -lt 12; $i++) {
        $machine = $machines[$rand.Next(0, $machines.Count)]
        $sessions.Add([pscustomobject]@{
            SessionKey  = "sess-pre-$i"
            UserId      = "usr-$($rand.Next(0, $UserCount))"
            MachineId   = $machine.Id
            StartUtc    = $windowStart.AddDays(-1 * (2 + $rand.Next(0, 20)))
            EndUtc      = $null
            SessionType = 0
            IsAnonymous = $false
        })
    }

    $applications = @(
        [pscustomobject]@{ Id = 'app-0'; Name = 'Excel';    PublishedName = 'Microsoft Excel' }
        [pscustomobject]@{ Id = 'app-1'; Name = 'Outlook';  PublishedName = 'Microsoft Outlook' }
        [pscustomobject]@{ Id = 'app-2'; Name = 'SAPGUI';   PublishedName = 'SAP GUI' }
        [pscustomobject]@{ Id = 'app-3'; Name = 'Notepad';  PublishedName = 'Notepad' }
        [pscustomobject]@{ Id = 'app-4'; Name = 'Chrome';   PublishedName = 'Google Chrome' }
    )

    # List[object] with .Add(), not `$x = @(); $x += ...` in the loop below.
    # PowerShell's += on an array reallocates and copies the ENTIRE array on
    # every append, making the loop O(n^2) in session count -- exactly the
    # same footgun Sessions above avoids by accumulating into a List. This
    # was measured costing tens of seconds at a few hundred synthetic users
    # and minutes at a few hundred more; every -DemoData run and every
    # preview paid it. Matches the Sessions convention immediately above:
    # accumulate into a List, convert once with .ToArray() at the end.
    $appInstances = New-Object System.Collections.Generic.List[object]
    $connections = New-Object System.Collections.Generic.List[object]
    $clientVersions = @('24.2.0.65','23.11.0.82','24.5.1.13','2402.10')

    foreach ($s in $sessions) {
        # The -f argument lists below are parenthesized, unlike the identical
        # expressions when this object was built via plain `+=` assignment
        # (see git history): a bare comma-separated -f argument list parses
        # fine as a statement, but the parser reads it differently once the
        # whole hashtable literal becomes an argument to a method call
        # (.Add(...) is expression-mode, not statement-mode) -- without the
        # parens here the comma is misread as separating .Add()'s own
        # arguments instead of -f's, and the hashtable literal never closes.
        $connections.Add([pscustomobject]@{
            Id = "conn-$($s.SessionKey)"
            SessionKey = $s.SessionKey
            ClientName = ("EP-{0:D4}" -f $rand.Next(0, [math]::Max(1, [int]($UserCount * 1.2))))
            ClientAddress = ("10.{0}.{1}.{2}" -f $rand.Next(0,255), $rand.Next(0,255), $rand.Next(1,254))
            ClientVersion = $clientVersions[$rand.Next(0, $clientVersions.Count)]
            IsReconnect = $false
        })
        if ($s.SessionType -eq 1) {
            $n = 1 + $rand.Next(0, 3)
            for ($a = 0; $a -lt $n; $a++) {
                # No StartDate/EndDate: the live fetcher's ApplicationInstances
                # select list omits them (nothing downstream reads them, and
                # Get-CitrixLookupRecords never applies ConvertTo-UtcDateTime to
                # this entity the way Get-CitrixSessionRecords does for Sessions).
                # Adding them here but not there would make the two diverge in
                # shape - real [datetime] on demo data, raw ISO-8601 string on a
                # real site - a mismatch that would pass every test against demo
                # data and only break against a real customer. If a future task
                # needs application duration, add the fields to BOTH the live
                # -Fields list above AND here, with a ConvertTo-UtcDateTime call
                # on the live side to match Sessions' handling.
                $appInstances.Add([pscustomobject]@{
                    Id = "ai-$($s.SessionKey)-$a"
                    SessionKey = $s.SessionKey
                    ApplicationId = "app-$($rand.Next(0, $applications.Count))"
                })
            }
        }
    }

    [pscustomobject]@{
        # NOTE: @($sessions) on a List[object] throws "Argument types do not
        # match" (ArgumentException) on some PowerShell 5.1 builds - the same
        # DLR binder issue ODataClient.ps1 works around with .ToArray()
        # instead of @(). Keep this consistent with that convention.
        Sessions             = $sessions.ToArray()
        Users                = $users
        Machines             = $machines
        DesktopGroups        = $groups
        Catalogs             = $catalogs
        # Same DLR-binder note as Sessions above: .ToArray(), not @(), on a
        # List[object].
        Connections          = $connections.ToArray()
        ApplicationInstances = $appInstances.ToArray()
        Applications         = $applications
        # Synthetic data is never fetched, so it is never truncated.
        FetchWarnings        = @()
        IsDemo               = $true
        # Mirrors the live Invoke-AuditPreflight shape field for field,
        # including HistoryStartUtc and RetentionBasis, so the analysis stage
        # exercises exactly the same code path on demo data as on live data.
        Preflight            = [pscustomobject]@{
            Reachable             = $true
            OldestSessionUtc      = $windowStart.AddDays(-22)
            OldestEndedSessionUtc = $windowStart.AddDays(-22)
            HistoryStartUtc       = $windowStart.AddDays(-22)
            RetentionBasis        = 'OldestEndedSession'
            AvailableDays         = $totalDays + 22
            TruncatedWindows      = @()
            CheckedAtUtc          = $Config.RunStartUtc
        }
    }
}

# endregion 50-DataModel.ps1

# ----------------------------------------------------------------------------
# region 60-Analytics.ps1
# ----------------------------------------------------------------------------
# ============================================================================
#  Analytics
#
#  Pure functions over the normalised dataset. No network, no formatting, no
#  global state. Correctness matters more here than anywhere else in the
#  script, because a wrong number in this file becomes a wrong usage figure
#  in a customer's licensing conversation, and nothing downstream would
#  reveal it.
# ============================================================================

function Assert-ValidAnalyticsWindow {
    <#
    .SYNOPSIS
        Fails loudly on a reversed window instead of letting one produce a
        silent, misleadingly-empty result.
    .DESCRIPTION
        A window with WindowEndUtc earlier than WindowStartUtc can only come
        from an upstream date-arithmetic bug. Left unchecked it does not
        error -- it quietly reports zero concurrent sessions, which is
        exactly the silent-wrong-answer failure this file exists to avoid.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][datetime] $WindowStartUtc,
        [Parameter(Mandatory)][datetime] $WindowEndUtc
    )

    if ($WindowEndUtc -lt $WindowStartUtc) {
        throw "WindowEndUtc ($WindowEndUtc) is earlier than WindowStartUtc ($WindowStartUtc); the window is reversed."
    }
}

function Get-ClampedSessionInterval {
    <#
    .SYNOPSIS
        Reduces a session to the portion falling inside a window.
    .DESCRIPTION
        Returns $null when the session does not overlap the window at all.
        A session with no end date is still running and is treated as ending
        at the instant the run started.

        A session whose clamped interval has zero (or negative) length is
        also dropped, and so counts towards no figure in the report -- not
        TotalSessions, not UniqueUsers, not concurrency. That covers a
        session recorded as starting and ending at the same instant (a failed
        or instantly-abandoned launch), and one that touches the window only
        at its very first or very last instant. The interval is half-open,
        [start, end), throughout this file, so a zero-length interval
        genuinely contains no sampled instant and cannot occupy a licence for
        any measurable time. The alternative -- counting it as one session --
        would inflate TotalSessions with launches that never ran.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] $Session,
        [Parameter(Mandatory)][datetime] $WindowStartUtc,
        [Parameter(Mandatory)][datetime] $WindowEndUtc,
        [Parameter(Mandatory)][datetime] $NowUtc
    )

    Assert-ValidAnalyticsWindow -WindowStartUtc $WindowStartUtc -WindowEndUtc $WindowEndUtc

    $start = $Session.StartUtc
    if (-not $start) { return $null }

    # No end date means the session is open right now.
    $end = if ($Session.EndUtc) { $Session.EndUtc } else { $NowUtc }

    # A session cannot be counted past the moment we observed the site.
    if ($end -gt $NowUtc) { $end = $NowUtc }

    if ($start -lt $WindowStartUtc) { $start = $WindowStartUtc }
    if ($end   -gt $WindowEndUtc)   { $end   = $WindowEndUtc }

    if ($end -le $start) { return $null }

    [pscustomobject]@{ StartUtc = $start; EndUtc = $end }
}

function Get-PeakConcurrency {
    <#
    .SYNOPSIS
        Exact maximum concurrent sessions in a window, by sweep line.
    .DESCRIPTION
        Session intervals become +1 events at their start and -1 events at
        their end. Sorting by timestamp and sweeping a running counter yields
        the true maximum in a single pass, correct for any overlap pattern.

        End events sort before start events at an identical timestamp, so a
        session ending exactly as another begins is not counted as concurrent.
        Without that tie-break every shift changeover would inflate the peak.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()] $Sessions,
        [Parameter(Mandatory)][datetime] $WindowStartUtc,
        [Parameter(Mandatory)][datetime] $WindowEndUtc,
        [Parameter(Mandatory)][datetime] $NowUtc
    )

    Assert-ValidAnalyticsWindow -WindowStartUtc $WindowStartUtc -WindowEndUtc $WindowEndUtc

    $events = New-Object System.Collections.Generic.List[object]

    foreach ($s in $Sessions) {
        $i = Get-ClampedSessionInterval -Session $s -WindowStartUtc $WindowStartUtc `
            -WindowEndUtc $WindowEndUtc -NowUtc $NowUtc
        if (-not $i) { continue }
        # Order 0 = end, 1 = start, so ends are processed first on a tie.
        [void] $events.Add([pscustomobject]@{ At = $i.EndUtc;   Delta = -1; Order = 0 })
        [void] $events.Add([pscustomobject]@{ At = $i.StartUtc; Delta =  1; Order = 1 })
    }

    if ($events.Count -eq 0) {
        return [pscustomobject]@{ Peak = 0; PeakAtUtc = $null }
    }

    $sorted = $events | Sort-Object -Property At, Order

    $current = 0; $peak = 0; $peakAt = $null
    foreach ($e in $sorted) {
        $current += $e.Delta
        if ($current -gt $peak) { $peak = $current; $peakAt = $e.At }
    }

    [pscustomobject]@{ Peak = $peak; PeakAtUtc = $peakAt }
}

function Get-ConcurrencySeries {
    <#
    .SYNOPSIS
        Concurrent session count sampled at fixed intervals.
    .DESCRIPTION
        Feeds the trend chart and the percentile figures. A bucket's value is
        the number of sessions open at the instant the bucket starts.

        Buckets are filled by computing each session's bucket-index range
        once and incrementing that range in a counts array, rather than
        testing every session against every bucket, because a 90-day window
        at 15-minute resolution is 8,640 buckets and a large site has
        hundreds of thousands of sessions.

        Returns a plain array. Callers must wrap the call in @() -- a
        single-bucket window's output would otherwise collapse to a scalar
        as it crosses the pipeline (see Invoke-CitrixODataQuery for the
        same convention and why it beats a producer-side unary comma).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()] $Sessions,
        [Parameter(Mandatory)][datetime] $WindowStartUtc,
        [Parameter(Mandatory)][datetime] $WindowEndUtc,
        [Parameter(Mandatory)][datetime] $NowUtc,
        [int] $IntervalMinutes = 15
    )

    Assert-ValidAnalyticsWindow -WindowStartUtc $WindowStartUtc -WindowEndUtc $WindowEndUtc

    # Ceiling, not Floor: a window whose length doesn't divide evenly by the
    # interval still has a real sample instant at the start of the trailing
    # remainder (e.g. a 65-minute window at 60-minute resolution must sample
    # both minute 0 and minute 60). Floor would drop that instant, and with
    # it any session open only during the remainder, understating the p95
    # that a licence count gets sized against. When the window divides
    # evenly, Ceiling and Floor agree, so no existing behaviour changes.
    $bucketCount = [int][math]::Ceiling(($WindowEndUtc - $WindowStartUtc).TotalMinutes / $IntervalMinutes)
    if ($bucketCount -le 0) { return @() }

    $counts = New-Object 'int[]' $bucketCount

    foreach ($s in $Sessions) {
        $i = Get-ClampedSessionInterval -Session $s -WindowStartUtc $WindowStartUtc `
            -WindowEndUtc $WindowEndUtc -NowUtc $NowUtc
        if (-not $i) { continue }

        # The first bucket whose start instant falls at or after the session
        # start, through the last bucket that starts before the session ends.
        $from = [int][math]::Ceiling(($i.StartUtc - $WindowStartUtc).TotalMinutes / $IntervalMinutes)
        $to   = [int][math]::Ceiling(($i.EndUtc   - $WindowStartUtc).TotalMinutes / $IntervalMinutes) - 1

        if ($from -lt 0) { $from = 0 }
        if ($to -ge $bucketCount) { $to = $bucketCount - 1 }

        for ($b = $from; $b -le $to; $b++) { $counts[$b]++ }
    }

    $series = New-Object System.Collections.Generic.List[object]
    for ($b = 0; $b -lt $bucketCount; $b++) {
        $series.Add([pscustomobject]@{
            TimeUtc = $WindowStartUtc.AddMinutes($b * $IntervalMinutes)
            Count   = $counts[$b]
        })
    }

    # Plain return, not a unary-comma-wrapped one: `@(Get-ConcurrencySeries
    # ...)` at the call site is the codebase's standard way to get a
    # reliably correct count for zero, one, or many buckets. A comma-wrapped
    # return works for a *bare* call but nests when the caller applies that
    # same @() -- silently turning "8,640 buckets" into "1", which is far
    # worse than the scalar collapse it was meant to prevent.
    return $series.ToArray()
}

function Resolve-AuditTimeZone {
    <#
    .SYNOPSIS
        Resolves a Windows time zone id, degrading to UTC rather than throwing.
    .DESCRIPTION
        Business-hours concurrency is a working-day figure, so it has to be
        computed in the SITE's local time, not in UTC. Computing it in UTC
        works by accident for a UK site and fails silently everywhere else:
        for a UTC+10 site, 09:00-17:00 local is 23:00-07:00 UTC, so an
        08:00-18:00 UTC business-hours test matches none of it and the figure
        collapses to zero; for a US site the number stays plausible while
        being wrong, which is worse.

        An unknown time zone id must not sink the run -- the same degrade-to-
        UTC choice Format-LocalTime already makes for display. The caller
        resolves the zone ONCE per run and reuses it, so the warning below is
        emitted once rather than per bucket.
    #>
    [CmdletBinding()]
    [OutputType([System.TimeZoneInfo])]
    param([AllowNull()][AllowEmptyString()][string] $TimeZoneId)

    if ([string]::IsNullOrWhiteSpace($TimeZoneId)) { return [System.TimeZoneInfo]::Utc }

    try {
        return [System.TimeZoneInfo]::FindSystemTimeZoneById($TimeZoneId)
    } catch {
        Write-AuditLog -Level Warn -Message "Time zone '$TimeZoneId' is not known to this machine; business-hours figures are computed in UTC instead. Peak and percentile figures for all hours are unaffected."
        return [System.TimeZoneInfo]::Utc
    }
}

function Test-InBusinessHours {
    <#
    .SYNOPSIS
        Whether a UTC instant falls inside the configured working day, in the
        site's own time zone.
    .DESCRIPTION
        Both the hour test AND the weekday test are applied to the converted
        local time. A Sydney Monday 09:00 is Sunday 23:00 UTC: testing the day
        of week in UTC discards a whole working morning as "weekend".

        Daylight saving is handled by the conversion itself, so a site that
        moves between BST and GMT (or between AEST and AEDT) keeps the same
        local working day across the change.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory)][datetime] $Utc,
        [Parameter(Mandatory)][System.TimeZoneInfo] $TimeZone,
        [Parameter(Mandatory)][int] $StartHour,
        [Parameter(Mandatory)][int] $EndHour
    )

    # ConvertTimeFromUtc throws on a DateTimeKind of Local, and treats
    # Unspecified as UTC. Normalising here keeps a caller that built its
    # series from an Unspecified-kind window from silently shifting.
    $u = if ($Utc.Kind -eq [System.DateTimeKind]::Local) {
        $Utc.ToUniversalTime()
    } else {
        [datetime]::SpecifyKind($Utc, [System.DateTimeKind]::Utc)
    }

    $local = [System.TimeZoneInfo]::ConvertTimeFromUtc($u, $TimeZone)

    return ($local.DayOfWeek -ne [DayOfWeek]::Saturday -and
            $local.DayOfWeek -ne [DayOfWeek]::Sunday -and
            $local.Hour -ge $StartHour -and
            $local.Hour -lt $EndHour)
}

function Get-Percentile {
    <#
    .SYNOPSIS
        Linear-interpolated percentile of a numeric set.
    .DESCRIPTION
        p95 is used instead of the raw maximum because a raw maximum is
        frequently one anomalous spike; p95 reflects the concurrency the
        environment sustains rather than a single outlier.
    #>
    [CmdletBinding()]
    [OutputType([double])]
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()] $Values,
        [Parameter(Mandatory)][ValidateRange(0, 100)][double] $Percentile
    )

    $sorted = @($Values | Sort-Object)
    if ($sorted.Count -eq 0) { return 0 }
    if ($sorted.Count -eq 1) { return [double] $sorted[0] }

    $rank = ($Percentile / 100.0) * ($sorted.Count - 1)
    $lower = [int][math]::Floor($rank)
    $upper = [int][math]::Ceiling($rank)

    if ($lower -eq $upper) { return [double] $sorted[$lower] }

    $weight = $rank - $lower
    return [double] ($sorted[$lower] + $weight * ($sorted[$upper] - $sorted[$lower]))
}

function Get-UniqueUserStats {
    <#
    .SYNOPSIS
        Distinct users with at least one session overlapping the window.
    .DESCRIPTION
        Anonymous sessions are counted separately and excluded from the user
        total, because they do not consume a named user licence. Folding them
        in would overstate the licence requirement.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()] $Sessions,
        [Parameter(Mandatory)][datetime] $WindowStartUtc,
        [Parameter(Mandatory)][datetime] $WindowEndUtc,
        [Parameter(Mandatory)][datetime] $NowUtc
    )

    # Validated directly (not only via the per-session Get-ClampedSessionInterval
    # calls below) so a reversed window is caught even when $Sessions is empty,
    # matching Get-PeakConcurrency / Get-ConcurrencySeries.
    Assert-ValidAnalyticsWindow -WindowStartUtc $WindowStartUtc -WindowEndUtc $WindowEndUtc

    # Relies on the default ordinal-ignore-case comparer of a PowerShell
    # hashtable literal: Citrix usernames are case-insensitive in practice
    # (DOMAIN\alice, DOMAIN\Alice, and DOMAIN\ALICE are the same licensed
    # person), so folding case here is required, not incidental. If this is
    # ever swapped for a case-sensitive collection (e.g.
    # Dictionary[string,bool]), the same person starts counting as multiple
    # licensed users -- see the dedicated pinning test in
    # tests/Analytics.Users.Tests.ps1.
    $userIds = @{}
    $anonymous = 0
    $unattributed = 0
    $total = 0

    foreach ($s in $Sessions) {
        $i = Get-ClampedSessionInterval -Session $s -WindowStartUtc $WindowStartUtc `
            -WindowEndUtc $WindowEndUtc -NowUtc $NowUtc
        if (-not $i) { continue }

        $total++

        if ($s.IsAnonymous) { $anonymous++; continue }
        # A blank/null UserId that is NOT flagged IsAnonymous still happens in
        # real Citrix exports (de-provisioned accounts, broker glitches). It
        # must not be counted as a named user, but it must not vanish without
        # a trace either: every in-window session lands in exactly one of the
        # three buckets below, so attributed sessions + AnonymousSessions +
        # UnattributedSessions == TotalSessions. UniqueUsers is NOT a term in
        # that sum -- it counts distinct people, not sessions, and one person
        # routinely has many sessions in a window.
        if ([string]::IsNullOrWhiteSpace($s.UserId)) { $unattributed++; continue }

        $userIds[$s.UserId] = $true
    }

    [pscustomobject]@{
        UniqueUsers          = $userIds.Count
        UniqueUserIds        = @($userIds.Keys)
        AnonymousSessions    = $anonymous
        UnattributedSessions = $unattributed
        TotalSessions        = $total
    }
}

function Get-DailyTrend {
    <#
    .SYNOPSIS
        Unique users and peak concurrency for each day in the window.
    .DESCRIPTION
        A session spanning several days counts on every day it touches, since
        it occupies a licence on each of them.

        Returns a plain array, per the codebase convention: callers wrap the
        call in @() to get a reliable Count for zero, one, or many days. See
        Get-ConcurrencySeries for the same convention and why a producer-side
        unary comma (`return ,$array`) must never be used instead -- it nests
        under a caller's @() and silently corrupts every Count downstream.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()] $Sessions,
        [Parameter(Mandatory)][datetime] $WindowStartUtc,
        [Parameter(Mandatory)][datetime] $WindowEndUtc,
        [Parameter(Mandatory)][datetime] $NowUtc
    )

    # Validated up front so a genuinely reversed window always throws, even
    # when $Sessions is empty. A zero-day window (WindowStartUtc equal to
    # WindowEndUtc) is not reversed, so this does not trip on that case --
    # dayCount below is simply 0 and the function returns @() before any
    # per-session processing happens.
    Assert-ValidAnalyticsWindow -WindowStartUtc $WindowStartUtc -WindowEndUtc $WindowEndUtc

    $firstDay = $WindowStartUtc.Date

    # dayCount is measured from the bucket anchor (midnight on the window's
    # start day), not from WindowStartUtc itself. WindowStartUtc is almost
    # never midnight in a real run (the window is built from NowUtc), and
    # measuring dayCount from it instead of firstDay used to leave the
    # trailing partial day with no bucket at all -- sessions landing there
    # were silently dropped rather than counted or erroring. The first row
    # can now cover a partial day (only the part of it inside the window),
    # which is correct: the window starts mid-day, so that row legitimately
    # represents less than 24 hours.
    $dayCount = [int][math]::Ceiling(($WindowEndUtc - $firstDay).TotalDays)
    if ($dayCount -le 0) { return @() }

    # Bucket sessions by the days they touch first, so each day's analytics run
    # over a small slice instead of the whole set. On a 90-day window with
    # hundreds of thousands of sessions, re-scanning everything per day would
    # turn a fast report into a multi-minute one.
    $byDay = @{}
    for ($d = 0; $d -lt $dayCount; $d++) { $byDay[$d] = New-Object System.Collections.Generic.List[object] }

    foreach ($s in $Sessions) {
        $i = Get-ClampedSessionInterval -Session $s -WindowStartUtc $WindowStartUtc `
            -WindowEndUtc $WindowEndUtc -NowUtc $NowUtc
        if (-not $i) { continue }

        $from = [int][math]::Floor(($i.StartUtc.Date - $firstDay).TotalDays)
        $to   = [int][math]::Floor(($i.EndUtc.Date   - $firstDay).TotalDays)
        if ($from -lt 0) { $from = 0 }
        if ($from -ge $dayCount) { $from = $dayCount - 1 }
        if ($to -lt 0) { $to = 0 }
        if ($to -ge $dayCount) { $to = $dayCount - 1 }

        # Clamping from/to independently can, in principle, still leave a
        # mismatched pair; skip rather than let the loop below run backwards
        # or fault, so a future range error degrades to "not counted here".
        if ($from -le $to) {
            for ($d = $from; $d -le $to; $d++) { [void] $byDay[$d].Add($s) }
        }
    }

    $results = New-Object System.Collections.Generic.List[object]

    for ($d = 0; $d -lt $dayCount; $d++) {
        $dayStart = [datetime]::SpecifyKind($firstDay.AddDays($d), [System.DateTimeKind]::Utc)
        $dayEnd = $dayStart.AddDays(1)
        $slice = $byDay[$d]

        $users = Get-UniqueUserStats -Sessions $slice -WindowStartUtc $dayStart `
            -WindowEndUtc $dayEnd -NowUtc $NowUtc
        $peak = Get-PeakConcurrency -Sessions $slice -WindowStartUtc $dayStart `
            -WindowEndUtc $dayEnd -NowUtc $NowUtc

        $results.Add([pscustomobject]@{
            DateUtc        = $dayStart
            UniqueUsers    = $users.UniqueUsers
            PeakConcurrent = $peak.Peak
            TotalSessions  = $users.TotalSessions
        })
    }

    # Plain return, not a unary-comma-wrapped one -- see the .DESCRIPTION note
    # above and Get-ConcurrencySeries for why.
    return $results.ToArray()
}

function Get-DeliveryGroupBreakdown {
    <#
    .SYNOPSIS
        Unique users and peak concurrency per delivery group.
    .DESCRIPTION
        Sessions are attributed through their machine. A session whose machine
        no longer exists is grouped under "Unknown" rather than discarded: the
        user was real, and dropping it would undercount the site.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][pscustomobject] $Dataset,
        [Parameter(Mandatory)][datetime] $WindowStartUtc,
        [Parameter(Mandatory)][datetime] $WindowEndUtc,
        [Parameter(Mandatory)][datetime] $NowUtc
    )

    # Validated unconditionally, up front -- matching the precedent
    # Get-UniqueUserStats set: a reversed window must throw immediately, not
    # quietly fall through to a zero-group result once every session is
    # attributed and filtered out downstream.
    Assert-ValidAnalyticsWindow -WindowStartUtc $WindowStartUtc -WindowEndUtc $WindowEndUtc

    $machineToGroup = @{}
    foreach ($m in $Dataset.Machines) { $machineToGroup[[string]$m.Id] = [string]$m.DesktopGroupId }

    $groupNames = @{}
    foreach ($g in $Dataset.DesktopGroups) { $groupNames[[string]$g.Id] = [string]$g.Name }

    $byGroup = @{}
    foreach ($s in $Dataset.Sessions) {
        $gid = $machineToGroup[[string]$s.MachineId]
        $name = if ($gid -and $groupNames.ContainsKey($gid)) { $groupNames[$gid] } else { 'Unknown (machine no longer present)' }
        if (-not $byGroup.ContainsKey($name)) {
            $byGroup[$name] = New-Object System.Collections.Generic.List[object]
        }
        [void] $byGroup[$name].Add($s)
    }

    $results = foreach ($name in $byGroup.Keys) {
        $slice = $byGroup[$name]
        $users = Get-UniqueUserStats -Sessions $slice -WindowStartUtc $WindowStartUtc `
            -WindowEndUtc $WindowEndUtc -NowUtc $NowUtc
        if ($users.TotalSessions -eq 0) { continue }
        $peak = Get-PeakConcurrency -Sessions $slice -WindowStartUtc $WindowStartUtc `
            -WindowEndUtc $WindowEndUtc -NowUtc $NowUtc

        [pscustomobject]@{
            Name           = $name
            UniqueUsers    = $users.UniqueUsers
            PeakConcurrent = $peak.Peak
            TotalSessions  = $users.TotalSessions
        }
    }

    @($results | Sort-Object -Property UniqueUsers -Descending)
}

function Get-SessionTypeBreakdown {
    <#
    .SYNOPSIS
        Desktop versus published application usage.
    .DESCRIPTION
        SessionType 0 is a desktop and 1 is an application. A user appearing in
        both is counted in both, since they used both resource types.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()] $Sessions,
        [Parameter(Mandatory)][datetime] $WindowStartUtc,
        [Parameter(Mandatory)][datetime] $WindowEndUtc,
        [Parameter(Mandatory)][datetime] $NowUtc
    )

    $desktop = @($Sessions | Where-Object { $_.SessionType -eq 0 })
    $app     = @($Sessions | Where-Object { $_.SessionType -eq 1 })

    $d = Get-UniqueUserStats -Sessions $desktop -WindowStartUtc $WindowStartUtc -WindowEndUtc $WindowEndUtc -NowUtc $NowUtc
    $a = Get-UniqueUserStats -Sessions $app     -WindowStartUtc $WindowStartUtc -WindowEndUtc $WindowEndUtc -NowUtc $NowUtc

    [pscustomobject]@{
        DesktopUsers        = $d.UniqueUsers
        DesktopSessions     = $d.TotalSessions
        ApplicationUsers    = $a.UniqueUsers
        ApplicationSessions = $a.TotalSessions
    }
}

function Get-ApplicationBreakdown {
    <#
    .SYNOPSIS
        Unique users and launch counts per published application.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][pscustomobject] $Dataset,
        [Parameter(Mandatory)][datetime] $WindowStartUtc,
        [Parameter(Mandatory)][datetime] $WindowEndUtc,
        [Parameter(Mandatory)][datetime] $NowUtc,
        [int] $Top = 20
    )

    # Validated unconditionally, up front -- even though the early return just
    # below would otherwise let a reversed window fall through silently
    # whenever ApplicationInstances happens to be empty.
    Assert-ValidAnalyticsWindow -WindowStartUtc $WindowStartUtc -WindowEndUtc $WindowEndUtc

    if (-not $Dataset.ApplicationInstances -or $Dataset.ApplicationInstances.Count -eq 0) { return @() }

    # Only sessions inside the window contribute, so build the eligible set once.
    # IsAnonymous is carried alongside UserId (not folded into a bare UserId
    # map) so that Launches can still count every in-window launch -- an
    # anonymous session is a real launch -- while UniqueUsers, like every
    # other unique-user count in this file, excludes anonymous sessions from
    # the named-user tally even when IsAnonymous is (unusually) paired with a
    # non-blank UserId.
    $sessionInfo = @{}
    foreach ($s in $Dataset.Sessions) {
        $i = Get-ClampedSessionInterval -Session $s -WindowStartUtc $WindowStartUtc `
            -WindowEndUtc $WindowEndUtc -NowUtc $NowUtc
        if ($i) {
            $sessionInfo[[string]$s.SessionKey] = [pscustomobject]@{
                UserId      = [string]$s.UserId
                IsAnonymous = [bool]$s.IsAnonymous
            }
        }
    }

    $appNames = @{}
    foreach ($a in $Dataset.Applications) {
        $label = if ($a.PublishedName) { [string]$a.PublishedName } else { [string]$a.Name }
        $appNames[[string]$a.Id] = $label
    }

    $launches = @{}; $appUsers = @{}
    foreach ($ai in $Dataset.ApplicationInstances) {
        $key = [string]$ai.SessionKey
        if (-not $sessionInfo.ContainsKey($key)) { continue }

        $appId = [string]$ai.ApplicationId
        $name = if ($appNames.ContainsKey($appId)) { $appNames[$appId] } else { "Unknown ($appId)" }

        if (-not $launches.ContainsKey($name)) {
            $launches[$name] = 0
            $appUsers[$name] = @{}
        }
        $launches[$name]++
        $info = $sessionInfo[$key]
        if ($info.UserId -and -not $info.IsAnonymous) { $appUsers[$name][$info.UserId] = $true }
    }

    $results = foreach ($name in $launches.Keys) {
        [pscustomobject]@{
            Name        = $name
            UniqueUsers = $appUsers[$name].Count
            Launches    = $launches[$name]
        }
    }

    @($results | Sort-Object -Property UniqueUsers, Launches -Descending | Select-Object -First $Top)
}

function Get-ClientDeviceBreakdown {
    <#
    .SYNOPSIS
        Distinct client devices, addresses and Workspace app versions.
    .DESCRIPTION
        Supports device-based licensing models, where the device count rather
        than the user count drives the contract.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][pscustomobject] $Dataset,
        [Parameter(Mandatory)][datetime] $WindowStartUtc,
        [Parameter(Mandatory)][datetime] $WindowEndUtc,
        [Parameter(Mandatory)][datetime] $NowUtc
    )

    # Validated unconditionally, up front -- even though the early return just
    # below would otherwise let a reversed window fall through silently
    # whenever Connections happens to be empty.
    Assert-ValidAnalyticsWindow -WindowStartUtc $WindowStartUtc -WindowEndUtc $WindowEndUtc

    if (-not $Dataset.Connections -or $Dataset.Connections.Count -eq 0) {
        return [pscustomobject]@{ UniqueDevices = 0; UniqueAddresses = 0; VersionCounts = @() }
    }

    $eligible = @{}
    foreach ($s in $Dataset.Sessions) {
        $i = Get-ClampedSessionInterval -Session $s -WindowStartUtc $WindowStartUtc `
            -WindowEndUtc $WindowEndUtc -NowUtc $NowUtc
        if ($i) { $eligible[[string]$s.SessionKey] = $true }
    }

    $devices = @{}; $addresses = @{}; $versions = @{}
    foreach ($c in $Dataset.Connections) {
        if (-not $eligible.ContainsKey([string]$c.SessionKey)) { continue }
        if ($c.ClientName)    { $devices[[string]$c.ClientName] = $true }
        if ($c.ClientAddress) { $addresses[[string]$c.ClientAddress] = $true }
        if ($c.ClientVersion) {
            $v = [string]$c.ClientVersion
            if (-not $versions.ContainsKey($v)) { $versions[$v] = 0 }
            $versions[$v]++
        }
    }

    $versionList = foreach ($v in $versions.Keys) {
        [pscustomobject]@{ Version = $v; Count = $versions[$v] }
    }

    [pscustomobject]@{
        UniqueDevices   = $devices.Count
        UniqueAddresses = $addresses.Count
        VersionCounts   = @($versionList | Sort-Object -Property Count -Descending)
    }
}

function Invoke-AuditAnalysis {
    <#
    .SYNOPSIS
        Runs every enabled analysis for every requested window.
    .OUTPUTS
        The analysis object consumed by the HTML report and the data export.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][pscustomobject] $Dataset,
        [Parameter(Mandatory)][pscustomobject] $Config
    )

    $now = $Config.RunStartUtc
    $available = if ($Dataset.Preflight) { [int] $Dataset.Preflight.AvailableDays } else { [int]::MaxValue }

    # The instant from which the site's history is actually present.
    # HistoryStartUtc is what Invoke-AuditPreflight derives from the oldest
    # ENDED session (the only data grooming can remove); OldestSessionUtc --
    # the oldest session START, which a single months-old still-running
    # desktop can set on its own -- is the fallback for a Preflight object
    # built before that field existed (older datasets, and test doubles that
    # construct a Preflight by hand).
    $historyStart = $null
    if ($Dataset.Preflight) {
        if ($Dataset.Preflight.PSObject.Properties['HistoryStartUtc'] -and $Dataset.Preflight.HistoryStartUtc) {
            $historyStart = $Dataset.Preflight.HistoryStartUtc
        } elseif ($Dataset.Preflight.OldestSessionUtc) {
            $historyStart = $Dataset.Preflight.OldestSessionUtc
        }
    }

    # Resolved once for the whole run: the warning inside it must not repeat
    # per window, and re-resolving per concurrency bucket would be measurable
    # on a 90-day window.
    $siteZone = Resolve-AuditTimeZone -TimeZoneId $Config.DisplayTimeZoneId

    $windows = foreach ($days in $Config.Days) {

        $start = $now.AddDays(-1 * $days)
        $isTruncated = $days -gt $available

        # When history is shorter than the requested window, analyse only the
        # span that actually exists. Analysing an empty stretch would drag the
        # percentiles toward zero and quietly flatter the customer's numbers.
        $effectiveStart = $start
        if ($isTruncated -and $historyStart) {
            $effectiveStart = $historyStart
        }

        # Defensive clamp: the history start is only ever expected to be at or
        # before $now, but it is computed by a prior, separate call
        # (Invoke-AuditPreflight) against whatever "now" that call used. A
        # dataset built against a different Config -- or simply stale data
        # reused across runs -- could carry a HistoryStartUtc later than
        # this run's $now, which would make effectiveStart later than $now
        # and trip Assert-ValidAnalyticsWindow's reversed-window guard below.
        # Clamping here keeps that a zero-length (not reversed) window
        # instead of a hard failure.
        if ($effectiveStart -gt $now) { $effectiveStart = $now }

        Write-AuditLog -Level Info -Message "Analysing the $days day window..."

        $users = Get-UniqueUserStats -Sessions $Dataset.Sessions -WindowStartUtc $effectiveStart `
            -WindowEndUtc $now -NowUtc $now
        $peak = Get-PeakConcurrency -Sessions $Dataset.Sessions -WindowStartUtc $effectiveStart `
            -WindowEndUtc $now -NowUtc $now
        $series = @(Get-ConcurrencySeries -Sessions $Dataset.Sessions -WindowStartUtc $effectiveStart `
            -WindowEndUtc $now -NowUtc $now -IntervalMinutes 15)

        $counts = @($series | ForEach-Object { $_.Count })

        # Business-hours concurrency, which is what most sizing conversations
        # actually care about. Weekend and overnight troughs otherwise drag the
        # percentiles down and understate the working-day requirement.
        #
        # Evaluated in the SITE's time zone ($siteZone, resolved once above),
        # never in UTC: see Test-InBusinessHours for why a UTC test silently
        # returns zero for an APAC site and a plausible-but-wrong number for a
        # US one. The zone is hoisted out of the pipeline deliberately -- a
        # 90-day window is 8,640 buckets and re-resolving it per bucket would
        # be measurable.
        $bizCounts = @(
            $series | Where-Object {
                Test-InBusinessHours -Utc $_.TimeUtc -TimeZone $siteZone `
                    -StartHour $Config.BusinessHourStart -EndHour $Config.BusinessHourEnd
            } | ForEach-Object { $_.Count }
        )

        $concurrency = [pscustomobject]@{
            Peak            = $peak.Peak
            PeakAtUtc       = $peak.PeakAtUtc
            P50             = Get-Percentile -Values $counts -Percentile 50
            P90             = Get-Percentile -Values $counts -Percentile 90
            P95             = Get-Percentile -Values $counts -Percentile 95
            P99             = Get-Percentile -Values $counts -Percentile 99
            BusinessP95     = Get-Percentile -Values $bizCounts -Percentile 95
            BusinessPeak    = $(if ($bizCounts.Count) { ($bizCounts | Measure-Object -Maximum).Maximum } else { 0 })
            # Carried on the result, not left implicit, so the report can state
            # what "business hours" actually means. A reader in another zone
            # otherwise assumes their own -- and the same figure means
            # different things at 08:00-18:00 London and 08:00-18:00 Sydney.
            BusinessHourStart  = [int] $Config.BusinessHourStart
            BusinessHourEnd    = [int] $Config.BusinessHourEnd
            BusinessTimeZoneId = $siteZone.Id
            Series          = $series
        }

        # Array-returning breakdowns are computed into plain local variables
        # BEFORE the [pscustomobject]@{} literal below, each with @() applied
        # directly to the call and nothing else -- never nested inside an
        # `if {} else {}` that itself sits inside a `$(...)` subexpression
        # assigned as a hashtable value. That nested shape looks equivalent
        # but is not: a single-element array returned from the "then" branch
        # crosses TWO separate implicit-output boundaries (the if-block's own
        # output, then the $(...) subexpression's capture of it), and each
        # boundary independently unrolls/collapses a lone element -- so a
        # one-delivery-group or one-application site (a small, entirely
        # ordinary customer) silently got a bare scalar with a $null .Count
        # instead of a one-element array, even with @() present at the call
        # site. A five-group demo dataset never has exactly one element in
        # any breakdown, so no existing test caught it. Verified empirically
        # (not just by inspection) before writing this comment. The one-line
        # `if ($Config.Include...) { $x = @(Get-... ...) }` form below has
        # only ONE implicit-output boundary -- the @() at the call site -- so
        # it does not suffer the same collapse.
        $deliveryGroups = $null
        if ($Config.IncludeDeliveryGroups) {
            $deliveryGroups = @(Get-DeliveryGroupBreakdown -Dataset $Dataset -WindowStartUtc $effectiveStart -WindowEndUtc $now -NowUtc $now)
        }

        $applications = $null
        if ($Config.IncludeApplications) {
            $applications = @(Get-ApplicationBreakdown -Dataset $Dataset -WindowStartUtc $effectiveStart -WindowEndUtc $now -NowUtc $now)
        }

        # Get-DailyTrend also returns a plain array and is exposed to the same
        # collapse risk in principle (a one-day window, or a window whose
        # trailing partial day is the only bucket) -- fixed here too, for
        # consistency and because the same collapse was proven to occur for
        # DeliveryGroups/Applications above using the structurally identical
        # nested-`$(if...)` shape this line previously used.
        $dailyTrend = $null
        if ($Config.IncludeTrend) {
            $dailyTrend = @(Get-DailyTrend -Sessions $Dataset.Sessions -WindowStartUtc $effectiveStart -WindowEndUtc $now -NowUtc $now)
        }

        [pscustomobject]@{
            Days               = $days
            WindowStartUtc     = $effectiveStart
            WindowEndUtc       = $now
            IsTruncated        = $isTruncated
            AvailableDays      = $available
            UniqueUsers        = $users.UniqueUsers
            UniqueUserIds      = $users.UniqueUserIds
            AnonymousSessions  = $users.AnonymousSessions
            UnattributedSessions = $users.UnattributedSessions
            TotalSessions      = $users.TotalSessions
            Concurrency        = $concurrency
            DeliveryGroups     = $deliveryGroups
            SessionTypes       = $(if ($Config.IncludeApplications) {
                Get-SessionTypeBreakdown -Sessions $Dataset.Sessions -WindowStartUtc $effectiveStart -WindowEndUtc $now -NowUtc $now
            } else { $null })
            Applications       = $applications
            ClientDevices      = $(if ($Config.IncludeClientDevices) {
                Get-ClientDeviceBreakdown -Dataset $Dataset -WindowStartUtc $effectiveStart -WindowEndUtc $now -NowUtc $now
            } else { $null })
            DailyTrend         = $dailyTrend
        }
    }

    # Assigned to a plain local BEFORE the object literal below, with @()
    # applied directly to the source and nothing else -- see the array-
    # convention note earlier in this function for why a nested
    # `$(if ...) { @(...) }` inside a hashtable value silently collapses a
    # one-element array to a scalar. A single fetch warning is by far the
    # likeliest case here.
    $fetchWarnings = @()
    if ($Dataset.PSObject.Properties['FetchWarnings']) {
        $fetchWarnings = @($Dataset.FetchWarnings)
    }

    [pscustomobject]@{
        GeneratedUtc     = $now
        Config           = $Config
        Preflight        = $Dataset.Preflight
        # Entities whose fetch came back knowingly incomplete or unbounded.
        # Surfaced onto the analysis so the report can warn next to the
        # figures they affect -- the same reason retention truncation gets an
        # in-report banner rather than only a log line.
        FetchWarnings    = $fetchWarnings
        Windows          = @($windows)
        SiteTotals       = [pscustomobject]@{
            DeliveryGroups = @($Dataset.DesktopGroups).Count
            Machines       = @($Dataset.Machines).Count
            KnownUsers     = @($Dataset.Users).Count
        }
        IsDemo           = [bool] $Dataset.IsDemo
    }
}

# endregion 60-Analytics.ps1

# ----------------------------------------------------------------------------
# region 65-Brand.ps1
# ----------------------------------------------------------------------------
# ============================================================================
#  Brand
#
#  The Citrix 2026 v1.2 brand system, as tokens. Values are documented in
#  docs/BRAND.md; that file is the authority and this one is its code form.
#
#  Fonts and logos are substituted in at build time as base64, so the
#  distributed script carries them and the generated report opens on an
#  air-gapped machine with no external request of any kind.
# ============================================================================

$script:BrandAssets = @{
    'font:FunnelDisplay' = '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'
    'font:PublicSans' = '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'
    'logo:black' = @'
<svg width="660" height="206" viewBox="0 0 660 206" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" fill="none" overflow="hidden"><style>
.MsftOfcThm_Text1_Fill_v2 {
 fill:#000000; 
}
</style>
<path d="M0 126.786C0 80.5185 31.0728 48.3027 75.372 48.3027 105.283 48.3027 129.624 63.2507 139.286 86.3712 140.474 88.7175 140.765 90.7474 140.765 93.1201 140.765 98.3928 136.646 101.899 131.393 101.899 127.274 101.899 123.79 99.5527 122.021 94.8601 114.391 76.1159 97.0992 65.2806 75.3984 65.2806 42.2664 65.2806 19.1136 90.4574 19.1136 126.786 19.1136 163.114 41.976 187.975 75.3984 187.975 97.0992 187.975 113.52 177.43 122.311 155.469 124.054 151.066 127.01 148.43 131.393 148.43 136.673 148.43 140.765 152.516 140.765 157.789 140.765 159.845 140.474 161.295 139.603 163.932 130.231 189.108 106.181 205.216 75.3984 205.216 30.8352 205.216 0.0264 173.001 0.0264 126.76" class="MsftOfcThm_Text1_Fill_v2" stroke="none" stroke-width="1" stroke-linecap="butt" stroke-linejoin="miter" stroke-miterlimit="4" fill="#000000" fill-opacity="1"/><path d="M251.514 192.377C251.514 198.81 248.003 202.923 241.851 202.923 235.7 202.923 232.479 198.81 232.479 192.377L232.479 10.7616C232.479 4.03897 235.991 0.216309 242.142 0.216309 248.293 0.216309 251.514 4.03897 251.514 10.7616L251.514 52.099 284.329 52.099C290.19 52.099 293.701 55.3153 293.701 60.588 293.701 65.8606 290.19 69.0769 284.329 69.0769L251.514 69.0769 251.514 192.351 251.514 192.377Z" class="MsftOfcThm_Text1_Fill_v2" stroke="none" stroke-width="1" stroke-linecap="butt" stroke-linejoin="miter" stroke-miterlimit="4" fill="#000000" fill-opacity="1"/><path d="M398.376 62.3543C398.376 67.6269 394.548 71.4232 389.004 71.4232 386.945 71.4232 384.886 70.8432 381.955 69.3669 377.573 67.3106 372.874 66.1506 366.722 66.1506 346.5 66.1506 333.3 83.7085 333.3 105.985L333.3 192.351C333.3 199.074 330.079 202.896 323.928 202.896 317.777 202.896 314.239 199.074 314.239 192.351L314.239 61.1943C314.239 54.4717 317.75 50.6491 323.928 50.6491 330.106 50.6491 333.3 54.4717 333.3 61.1943L333.3 64.9906C341.51 54.1554 354.394 48.3027 368.782 48.3027 379.922 48.3027 387.842 50.6491 393.413 54.1554 396.634 56.2117 398.402 58.848 398.402 62.3543" class="MsftOfcThm_Text1_Fill_v2" stroke="none" stroke-width="1" stroke-linecap="butt" stroke-linejoin="miter" stroke-miterlimit="4" fill="#000000" fill-opacity="1"/><path d="M417.702 61.1944C417.702 54.4718 421.213 50.6492 427.365 50.6492 433.516 50.6492 436.737 54.4718 436.737 61.1944L436.737 192.378C436.737 198.81 433.225 202.923 427.074 202.923 420.923 202.923 417.702 198.837 417.702 192.378L417.702 61.1944Z" class="MsftOfcThm_Text1_Fill_v2" stroke="none" stroke-width="1" stroke-linecap="butt" stroke-linejoin="miter" stroke-miterlimit="4" fill="#000000" fill-opacity="1"/><path d="M174.214 61.1944C174.214 54.4718 177.725 50.6492 183.876 50.6492 190.027 50.6492 193.248 54.4718 193.248 61.1944L193.248 192.378C193.248 198.81 189.737 202.923 183.586 202.923 177.435 202.923 174.214 198.837 174.214 192.378L174.214 61.1944Z" class="MsftOfcThm_Text1_Fill_v2" stroke="none" stroke-width="1" stroke-linecap="butt" stroke-linejoin="miter" stroke-miterlimit="4" fill="#000000" fill-opacity="1"/><path d="M183.508 0.216309C175.112 0.216309 168.776 6.5962 168.776 15.0588 168.776 23.5214 175.244 30.1385 183.508 30.1385 191.771 30.1385 198.714 23.5214 198.714 15.0588 198.714 6.5962 192.167 0.216309 183.508 0.216309Z" class="MsftOfcThm_Text1_Fill_v2" stroke="none" stroke-width="1" stroke-linecap="butt" stroke-linejoin="miter" stroke-miterlimit="4" fill="#000000" fill-opacity="1"/><path d="M535.868 138.491 483.966 199.1C481.907 201.736 479.267 202.923 476.336 202.923 471.056 202.923 467.255 198.836 467.255 194.144 467.255 192.087 467.836 189.741 469.895 187.421L524.147 126.232 471.954 66.4931C469.895 64.1468 468.733 62.0905 468.733 59.4542 468.733 54.7615 472.535 50.6752 478.105 50.6752 481.036 50.6752 483.095 51.8616 485.444 54.4979L535.868 113.63 586.292 54.4979C588.642 51.8616 590.701 50.6752 593.632 50.6752 599.202 50.6752 603.004 54.7879 603.004 59.4542 603.004 62.0905 601.842 64.1468 599.783 66.4931L547.59 126.232 601.842 187.421C603.901 189.767 604.482 192.114 604.482 194.144 604.482 198.836 600.68 202.923 595.4 202.923 592.47 202.923 589.83 201.763 587.771 199.1L535.868 138.491Z" class="MsftOfcThm_Text1_Fill_v2" stroke="none" stroke-width="1" stroke-linecap="butt" stroke-linejoin="miter" stroke-miterlimit="4" fill="#000000" fill-opacity="1"/><path d="M535.628 0.216309C527.232 0.216309 520.896 6.5962 520.896 15.0588 520.896 23.5214 527.364 30.1385 535.628 30.1385 543.891 30.1385 550.834 23.5214 550.834 15.0588 550.834 6.5962 544.313 0.216309 535.628 0.216309Z" class="MsftOfcThm_Text1_Fill_v2" stroke="none" stroke-width="1" stroke-linecap="butt" stroke-linejoin="miter" stroke-miterlimit="4" fill="#000000" fill-opacity="1"/><path d="M641.415 181.542 646.246 181.542C648.807 181.542 650.1 182.649 650.1 184.864 650.1 187.078 648.807 188.186 646.246 188.238L641.415 188.238 641.415 181.542ZM651.948 184.864C651.948 183.335 651.447 182.175 650.443 181.305 649.44 180.461 648.041 180.039 646.246 180.039L639.54 180.039 639.54 196.253 641.415 196.253 641.415 189.688 645.797 189.688 650.1 196.253 652.239 196.253 647.777 189.557C649.097 189.319 650.127 188.818 650.839 188.001 651.579 187.21 651.948 186.156 651.948 184.864ZM645.005 201.156C642.391 201.156 640.121 200.603 638.167 199.548 636.214 198.467 634.683 196.938 633.627 194.935 632.571 192.957 632.016 190.664 632.016 188.054 632.016 185.444 632.544 183.15 633.627 181.173 634.683 179.196 636.214 177.64 638.167 176.56 640.121 175.479 642.418 174.951 645.005 174.951 647.592 174.951 649.889 175.505 651.843 176.56 653.796 177.64 655.327 179.169 656.41 181.173 657.492 183.15 658.02 185.444 658.02 188.054 658.02 190.664 657.492 192.957 656.41 194.935 655.327 196.912 653.796 198.467 651.843 199.548 649.889 200.629 647.592 201.156 645.005 201.156ZM645.005 173.159C642.022 173.159 639.408 173.765 637.138 175.004 634.867 176.217 633.125 177.983 631.884 180.25 630.643 182.518 630.01 185.128 630.01 188.08 630.01 191.033 630.617 193.643 631.884 195.91 633.125 198.177 634.894 199.917 637.138 201.156 639.408 202.395 642.022 203.002 645.005 203.002 647.988 203.002 650.602 202.395 652.872 201.156 655.143 199.944 656.885 198.204 658.126 195.936 659.367 193.696 660 191.059 660 188.08 660 185.101 659.393 182.491 658.126 180.25 656.885 178.01 655.116 176.27 652.872 175.03 650.602 173.791 647.988 173.185 645.005 173.185" class="MsftOfcThm_Text1_Fill_v2" stroke="none" stroke-width="1" stroke-linecap="butt" stroke-linejoin="miter" stroke-miterlimit="4" fill="#000000" fill-opacity="1"/></svg>
'@
    'logo:white' = @'
<svg width="660" height="206" viewBox="0 0 660 206" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" fill="none" overflow="hidden"><style>
.MsftOfcThm_Background1_Fill_v2 {
 fill:#FFFFFF; 
}
</style>
<path d="M0 126.786C0 80.5185 31.0728 48.3027 75.372 48.3027 105.283 48.3027 129.624 63.2507 139.286 86.3712 140.474 88.7175 140.765 90.7474 140.765 93.1201 140.765 98.3928 136.646 101.899 131.393 101.899 127.274 101.899 123.79 99.5527 122.021 94.8601 114.391 76.1159 97.0992 65.2806 75.3984 65.2806 42.2664 65.2806 19.1136 90.4574 19.1136 126.786 19.1136 163.114 41.976 187.975 75.3984 187.975 97.0992 187.975 113.52 177.43 122.311 155.469 124.054 151.066 127.01 148.43 131.393 148.43 136.673 148.43 140.765 152.516 140.765 157.789 140.765 159.845 140.474 161.295 139.603 163.932 130.231 189.108 106.181 205.216 75.3984 205.216 30.8352 205.216 0.0264 173.001 0.0264 126.76" class="MsftOfcThm_Background1_Fill_v2" fill="#FFFFFF"/><path d="M251.514 192.377C251.514 198.81 248.003 202.923 241.851 202.923 235.7 202.923 232.479 198.81 232.479 192.377L232.479 10.7616C232.479 4.03897 235.991 0.216309 242.142 0.216309 248.293 0.216309 251.514 4.03897 251.514 10.7616L251.514 52.099 284.329 52.099C290.19 52.099 293.701 55.3153 293.701 60.588 293.701 65.8606 290.19 69.0769 284.329 69.0769L251.514 69.0769 251.514 192.351 251.514 192.377Z" class="MsftOfcThm_Background1_Fill_v2" fill="#FFFFFF"/><path d="M398.376 62.3543C398.376 67.6269 394.548 71.4232 389.004 71.4232 386.945 71.4232 384.886 70.8432 381.955 69.3669 377.573 67.3106 372.874 66.1506 366.722 66.1506 346.5 66.1506 333.3 83.7085 333.3 105.985L333.3 192.351C333.3 199.074 330.079 202.896 323.928 202.896 317.777 202.896 314.239 199.074 314.239 192.351L314.239 61.1943C314.239 54.4717 317.75 50.6491 323.928 50.6491 330.106 50.6491 333.3 54.4717 333.3 61.1943L333.3 64.9906C341.51 54.1554 354.394 48.3027 368.782 48.3027 379.922 48.3027 387.842 50.6491 393.413 54.1554 396.634 56.2117 398.402 58.848 398.402 62.3543" class="MsftOfcThm_Background1_Fill_v2" fill="#FFFFFF"/><path d="M417.702 61.1944C417.702 54.4718 421.213 50.6492 427.365 50.6492 433.516 50.6492 436.737 54.4718 436.737 61.1944L436.737 192.378C436.737 198.81 433.225 202.923 427.074 202.923 420.923 202.923 417.702 198.837 417.702 192.378L417.702 61.1944Z" class="MsftOfcThm_Background1_Fill_v2" fill="#FFFFFF"/><path d="M174.214 61.1944C174.214 54.4718 177.725 50.6492 183.876 50.6492 190.027 50.6492 193.248 54.4718 193.248 61.1944L193.248 192.378C193.248 198.81 189.737 202.923 183.586 202.923 177.435 202.923 174.214 198.837 174.214 192.378L174.214 61.1944Z" class="MsftOfcThm_Background1_Fill_v2" fill="#FFFFFF"/><path d="M183.508 0.216309C175.112 0.216309 168.776 6.5962 168.776 15.0588 168.776 23.5214 175.244 30.1385 183.508 30.1385 191.771 30.1385 198.714 23.5214 198.714 15.0588 198.714 6.5962 192.167 0.216309 183.508 0.216309Z" class="MsftOfcThm_Background1_Fill_v2" fill="#FFFFFF"/><path d="M535.868 138.491 483.966 199.1C481.907 201.736 479.267 202.923 476.336 202.923 471.056 202.923 467.255 198.836 467.255 194.144 467.255 192.087 467.836 189.741 469.895 187.421L524.147 126.232 471.954 66.4931C469.895 64.1468 468.733 62.0905 468.733 59.4542 468.733 54.7615 472.535 50.6752 478.105 50.6752 481.036 50.6752 483.095 51.8616 485.444 54.4979L535.868 113.63 586.292 54.4979C588.642 51.8616 590.701 50.6752 593.632 50.6752 599.202 50.6752 603.004 54.7879 603.004 59.4542 603.004 62.0905 601.842 64.1468 599.783 66.4931L547.59 126.232 601.842 187.421C603.901 189.767 604.482 192.114 604.482 194.144 604.482 198.836 600.68 202.923 595.4 202.923 592.47 202.923 589.83 201.763 587.771 199.1L535.868 138.491Z" class="MsftOfcThm_Background1_Fill_v2" fill="#FFFFFF"/><path d="M535.628 0.216309C527.232 0.216309 520.896 6.5962 520.896 15.0588 520.896 23.5214 527.364 30.1385 535.628 30.1385 543.891 30.1385 550.834 23.5214 550.834 15.0588 550.834 6.5962 544.313 0.216309 535.628 0.216309Z" class="MsftOfcThm_Background1_Fill_v2" fill="#FFFFFF"/><path d="M641.415 181.542 646.246 181.542C648.807 181.542 650.1 182.649 650.1 184.864 650.1 187.078 648.807 188.186 646.246 188.238L641.415 188.238 641.415 181.542ZM651.948 184.864C651.948 183.335 651.447 182.175 650.443 181.305 649.44 180.461 648.041 180.039 646.246 180.039L639.54 180.039 639.54 196.253 641.415 196.253 641.415 189.688 645.797 189.688 650.1 196.253 652.239 196.253 647.777 189.557C649.097 189.319 650.127 188.818 650.839 188.001 651.579 187.21 651.948 186.156 651.948 184.864ZM645.005 201.156C642.391 201.156 640.121 200.603 638.167 199.548 636.214 198.467 634.683 196.938 633.627 194.935 632.571 192.957 632.016 190.664 632.016 188.054 632.016 185.444 632.544 183.15 633.627 181.173 634.683 179.196 636.214 177.64 638.167 176.56 640.121 175.479 642.418 174.951 645.005 174.951 647.592 174.951 649.889 175.505 651.843 176.56 653.796 177.64 655.327 179.169 656.41 181.173 657.492 183.15 658.02 185.444 658.02 188.054 658.02 190.664 657.492 192.957 656.41 194.935 655.327 196.912 653.796 198.467 651.843 199.548 649.889 200.629 647.592 201.156 645.005 201.156ZM645.005 173.159C642.022 173.159 639.408 173.765 637.138 175.004 634.867 176.217 633.125 177.983 631.884 180.25 630.643 182.518 630.01 185.128 630.01 188.08 630.01 191.033 630.617 193.643 631.884 195.91 633.125 198.177 634.894 199.917 637.138 201.156 639.408 202.395 642.022 203.002 645.005 203.002 647.988 203.002 650.602 202.395 652.872 201.156 655.143 199.944 656.885 198.204 658.126 195.936 659.367 193.696 660 191.059 660 188.08 660 185.101 659.393 182.491 658.126 180.25 656.885 178.01 655.116 176.27 652.872 175.03 650.602 173.791 647.988 173.185 645.005 173.185" class="MsftOfcThm_Background1_Fill_v2" fill="#FFFFFF"/></svg>
'@
}

function Get-BrandPalette {
    <#
    .SYNOPSIS
        The Citrix brand palette as token name to hex.
    .DESCRIPTION
        Cobalt 70 is #0B2168. The printed style guide labels it #00236E, but
        its own RGB value (11/33/104) and the PowerPoint theme's dk2 both give
        #0B2168, so the printed hex is treated as a typo. See docs/BRAND.md.
    #>
    [CmdletBinding()]
    [OutputType([hashtable])]
    param()

    @{
        Cobalt50  = '#0045DB'   # Primary. Anchors the visual identity.
        Cobalt60  = '#0033A3'   # Derived tint, no official value published.
        Cobalt70  = '#0B2168'
        Mint20    = '#C2FFD0'
        Mint40    = '#71FD95'
        Slate10   = '#C4D6DA'   # Derived tint, no official value published.
        Slate20   = '#98B8BE'
        Jasper80  = '#2E5957'
        Jasper100 = '#172B29'
        Aether10  = '#EFF8FF'
        Aether20  = '#D6EDFF'
        Canvas20  = '#F3EEE8'
        Canvas30  = '#E5DDD2'   # Derived tint, no official value published.
    }
}

function ConvertFrom-HexColor {
    <#
    .SYNOPSIS
        Parses #RRGGBB into r, g, b components.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string] $Hex)

    $h = $Hex.TrimStart('#')
    [pscustomobject]@{
        R = [Convert]::ToInt32($h.Substring(0, 2), 16)
        G = [Convert]::ToInt32($h.Substring(2, 2), 16)
        B = [Convert]::ToInt32($h.Substring(4, 2), 16)
    }
}

function Get-RelativeLuminance {
    <#
    .SYNOPSIS
        WCAG relative luminance of a colour.
    #>
    [CmdletBinding()]
    [OutputType([double])]
    param([Parameter(Mandatory)][string] $Hex)

    $c = ConvertFrom-HexColor -Hex $Hex
    $channels = @($c.R, $c.G, $c.B) | ForEach-Object {
        $v = $_ / 255.0
        if ($v -le 0.03928) { $v / 12.92 } else { [math]::Pow((($v + 0.055) / 1.055), 2.4) }
    }
    return 0.2126 * $channels[0] + 0.7152 * $channels[1] + 0.0722 * $channels[2]
}

function Get-ContrastRatio {
    <#
    .SYNOPSIS
        WCAG contrast ratio between two colours, from 1 to 21.
    .DESCRIPTION
        The brand requires AA at minimum: 4.5:1 for normal text, 3:1 for large.
        The test suite asserts every pairing the report uses, so a later palette
        edit cannot silently drop the report below AA.
    #>
    [CmdletBinding()]
    [OutputType([double])]
    param(
        [Parameter(Mandatory)][string] $Foreground,
        [Parameter(Mandatory)][string] $Background
    )

    $l1 = Get-RelativeLuminance -Hex $Foreground
    $l2 = Get-RelativeLuminance -Hex $Background
    $lighter = [math]::Max($l1, $l2)
    $darker  = [math]::Min($l1, $l2)
    return ($lighter + 0.05) / ($darker + 0.05)
}

function Get-BrandFontCss {
    <#
    .SYNOPSIS
        @font-face rules with base64 WOFF2 sources.
    .DESCRIPTION
        Returns an empty string when no fonts were vendored, in which case the
        report falls back to its system font stack. Missing fonts degrade the
        look; they never break the report.

        Both families are vendored as a single variable-font file that covers
        every weight the brand uses (tools/Get-BrandFonts.ps1 vendors exactly
        one .woff2 per family, since every requested weight instance resolves
        to the identical underlying file on Google Fonts). So each family
        emits exactly one @font-face declaring a weight RANGE - not one block
        per discrete weight - and the browser interpolates the font's own
        wght axis to render intermediate weights.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param()

    if (-not $script:BrandAssets) { return '' }

    $families = @{
        'FunnelDisplay' = @{ Name = 'Funnel Display'; Weight = '500 600' }
        'PublicSans'    = @{ Name = 'Public Sans';    Weight = '300 600' }
    }

    $rules = New-Object System.Collections.Generic.List[string]
    foreach ($key in $script:BrandAssets.Keys) {
        if (-not $key.StartsWith('font:')) { continue }
        $familyKey = $key.Substring(5)
        $family = $families[$familyKey]
        if (-not $family) { continue }

        [void] $rules.Add(@"
@font-face {
  font-family: '$($family.Name)';
  font-style: normal;
  font-weight: $($family.Weight);
  font-display: swap;
  src: url(data:font/woff2;base64,$($script:BrandAssets[$key])) format('woff2');
}
"@)
    }
    return ($rules -join "`n")
}

function Get-BrandLogoSvg {
    <#
    .SYNOPSIS
        The Citrix wordmark as inline SVG.
    .DESCRIPTION
        The distributed single-file script carries the wordmark as base64 via
        $script:BrandAssets, substituted in at build time, because it has no
        side files to fall back on.

        When this file is dot-sourced directly from source (as the test suite
        does, and as $script:BrandAssets being unset always implies), the SVG
        is read straight from assets/brand instead. Those files are already
        committed to the repository, so this path never depends on the build
        or on network access.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([ValidateSet('black', 'white')][string] $Variant = 'black')

    if ($script:BrandAssets -and $script:BrandAssets.ContainsKey("logo:$Variant")) {
        return $script:BrandAssets["logo:$Variant"]
    }

    $svgPath = Join-Path (Split-Path -Parent $PSScriptRoot) "assets/brand/citrix-wordmark-$Variant.svg"
    if (Test-Path $svgPath) {
        return Get-Content -Path $svgPath -Raw
    }
    return ''
}

function Get-BrandCss {
    <#
    .SYNOPSIS
        The complete report stylesheet.
    .DESCRIPTION
        Page chrome uses the brand-preferred Canvas 20 | Cobalt 50 | Black
        pairing; the header band uses Jasper 80 | Mint 40 | White. Public Sans
        never appears above weight 600, per the brand's prohibition on Bold,
        ExtraBold and Black in that family.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param()

    $p = Get-BrandPalette
    $tokens = ($p.Keys | Sort-Object | ForEach-Object {
        "  --{0}: {1};" -f $_.ToLower(), $p[$_]
    }) -join "`n"

    $fontCss = Get-BrandFontCss

    @"
$fontCss

:root {
$tokens
  --page-bg: $($p['Canvas20']);
  --surface: #FFFFFF;
  --text: $($p['Jasper100']);
  --text-muted: $($p['Jasper80']);
  --accent: $($p['Cobalt50']);
  --accent-deep: $($p['Cobalt70']);
  --positive: $($p['Mint40']);
  --rule: $($p['Slate20']);

  /* Brand type ramp: 1.4x ratio, -2% letter spacing. */
  --font-head: 'Funnel Display', 'Segoe UI Semibold', 'Segoe UI', Arial, sans-serif;
  --font-body: 'Public Sans', 'Segoe UI', Arial, sans-serif;
}

* { box-sizing: border-box; }

body {
  margin: 0;
  background: var(--page-bg);
  color: var(--text);
  font-family: var(--font-body);
  font-weight: 400;
  font-size: 16px;
  line-height: 1.3;            /* Brand: 130% for body copy. */
  letter-spacing: -0.02em;     /* Brand: -2% letter spacing. */
}

h1, h2, h3 {
  font-family: var(--font-head);
  font-weight: 600;            /* Funnel Display SemiBold. */
  line-height: 1.1;            /* Brand: 110% for headers. */
  letter-spacing: -0.02em;
  margin: 0 0 0.5em 0;
}

h1 { font-size: 2.75rem; }
h2 { font-size: 1.96rem; }     /* 2.75 / 1.4 */
h3 { font-size: 1.4rem; }

p { margin: 0 0 1em 0; }

.wrap { max-width: 1180px; margin: 0 auto; padding: 0 32px 64px; }

/* Header band: Jasper 80 | Mint 40 | White - a brand-preferred pairing. */
.report-header {
  background: $($p['Jasper80']);
  color: #FFFFFF;
  padding: 40px 0 48px;
  margin-bottom: 40px;
}
.report-header .wrap { padding-bottom: 0; }
.report-header h1 { color: #FFFFFF; }
.report-header .sub { color: $($p['Mint40']); font-weight: 500; }
.report-header .logo { height: 34px; margin-bottom: 28px; }
.report-header .logo svg { height: 34px; width: auto; }

/* Logo/title on the left, print action on the right - the same .wrap
   centring and horizontal padding as the rest of the page, just laid out
   as a row instead of stacked. */
.header-row { display: flex; justify-content: space-between; align-items: flex-start; gap: 24px; flex-wrap: wrap; }
.header-actions { text-align: right; }
.print-btn {
  background: var(--accent);
  color: #FFFFFF;
  border: none;
  border-radius: 8px;
  padding: 12px 22px;
  font-family: var(--font-body);
  font-weight: 600;
  font-size: 0.95rem;
  letter-spacing: 0;
  cursor: pointer;
}
.print-btn:hover, .print-btn:focus { background: var(--accent-deep); }
.print-hint { color: $($p['Mint40']); font-size: 0.8rem; margin: 8px 0 0; }

.card {
  background: var(--surface);
  border: 1px solid var(--rule);
  border-radius: 10px;
  padding: 28px;
  margin-bottom: 28px;
}

.kpi-row { display: flex; flex-wrap: wrap; gap: 20px; margin-bottom: 28px; }
.kpi {
  flex: 1 1 210px;
  background: var(--surface);
  border: 1px solid var(--rule);
  border-left: 5px solid var(--accent);
  border-radius: 10px;
  padding: 22px 24px;
}
.kpi .label {
  font-size: 0.8rem; text-transform: uppercase; letter-spacing: 0.06em;
  color: var(--text-muted); font-weight: 500; margin-bottom: 8px;
}
.kpi .value { font-family: var(--font-head); font-size: 2.4rem; font-weight: 600; line-height: 1; }
.kpi .note { font-size: 0.82rem; color: var(--text-muted); margin-top: 8px; }

/* Truncation is called out where the figure appears, never in a footnote. */
.truncated {
  background: $($p['Aether20']);
  border-left: 5px solid var(--accent-deep);
  padding: 16px 20px; border-radius: 8px; margin: 0 0 24px;
  font-size: 0.92rem;
}
.truncated strong { font-weight: 600; }
.truncated a { color: var(--accent-deep); font-weight: 600; text-decoration: underline; }

/* Synthetic-data warning. Deliberately a separate class from .truncated: the
   two warnings are unrelated, and a demo report must never read as a real one. */
.notice-demo {
  background: $($p['Mint20']);
  border-left: 5px solid $($p['Jasper80']);
  padding: 16px 20px; border-radius: 8px; margin: 0 0 24px;
  font-size: 0.92rem;
}
.notice-demo strong { font-weight: 600; }

table { width: 100%; border-collapse: collapse; font-size: 0.94rem; }
th {
  text-align: left; font-weight: 600; color: var(--text-muted);
  border-bottom: 2px solid var(--rule); padding: 10px 12px;
  text-transform: uppercase; font-size: 0.76rem; letter-spacing: 0.06em;
}
td { padding: 11px 12px; border-bottom: 1px solid var(--rule); }
td.num, th.num { text-align: right; font-variant-numeric: tabular-nums; }
tbody tr:last-child td { border-bottom: none; }

.scroll-x { overflow-x: auto; }

.footnote { font-size: 0.85rem; color: var(--text-muted); }

@media print {
  @page { margin: 14mm 12mm; }

  body { background: #FFFFFF; }

  /* Preserve backgrounds and accent colours. Without this, a browser's
     default print behaviour is to drop background colours to save ink -
     a truncation banner that prints as plain white text on white loses
     the urgency it exists to carry, and the KPI accent border and header
     band go with it. Set broadly (inherited) rather than only on
     .report-header, so every banner and accent survives too. */
  body, .truncated, .notice-demo, .kpi {
    -webkit-print-color-adjust: exact;
    print-color-adjust: exact;
  }

  /* Screen-only chrome: the print button has no purpose on a printed page
     and a button printed onto a PDF looks broken. */
  .no-print { display: none !important; }

  .card, .kpi, .kpi-row { break-inside: avoid; border-color: #CCCCCC; }
  /* A card taller than a single page cannot honour break-inside: avoid -
     there is no page it fits on whole. Neither .card nor its children set
     overflow: hidden anywhere, so the degrade is the browser's normal
     fragmentation behaviour (the card starts on a fresh page and its
     content flows across as many pages as it needs), never clipping. */

  h1, h2, h3 { break-after: avoid; }
  p { orphans: 3; widows: 3; }

  /* A table that spans a page break is unreadable without its header
     repeating, and a data row split across the break is worse - half a
     row on each page. */
  thead { display: table-header-group; }
  tfoot { display: table-footer-group; }
  tbody tr { break-inside: avoid; }

  /* overflow-x: auto has no scrollbar on paper - left as auto, a table
     wider than the printed page would be clipped instead of scrollable. */
  .scroll-x { overflow-x: visible; }

  /* The truncation/retention/fetch-warning banner (.truncated) tells the
     reader a figure is a lower bound, not a true count. Orphaned onto a
     separate page from the figures it qualifies, the report becomes
     actively misleading, so it is kept intact and glued to whatever
     immediately follows it (the KPI row, in every case it appears). */
  .truncated, .notice-demo {
    break-inside: avoid;
    break-after: avoid-page;
  }
}
"@
}

# endregion 65-Brand.ps1

# ----------------------------------------------------------------------------
# region 70-Charts.ps1
# ----------------------------------------------------------------------------
# ============================================================================
#  SVG charts
#
#  Charts are generated as inline SVG rather than drawn by a JavaScript
#  library, so the report has no external dependency, opens on an air-gapped
#  machine, and prints correctly.
#
#  Colour is never used to carry category identity on its own here. Bar
#  charts (the common case: delivery groups, applications, users) draw every
#  bar in the same brand-primary hue - the label carries identity, the bar
#  length carries the value. The one chart that legitimately compares two
#  series side by side (New-SvgGroupedBarChart) uses exactly two hues,
#  Cobalt 50 and Jasper 80, chosen because both clear 3:1 contrast against
#  the white card the charts sit on and separate cleanly under deuteranopia
#  and tritanopia simulation - and it keeps per-bar value labels and a
#  legend as the secondary encoding that makes a two-hue comparison safe for
#  colour-blind readers. Three of the twelve palette tokens (Mint 40,
#  Slate 20, Aether 20) are too light to read as a data mark on a light
#  surface at all, even though they are legitimate brand background and
#  accent tints elsewhere in the stylesheet; Get-ChartSeriesColor excludes
#  them for that reason. See docs/BRAND.md and the Task 12 report for the
#  measured contrast ratios behind this.
# ============================================================================

# Minimum horizontal gap, in pixels, a New-SvgAreaChart x-axis label is
# allowed from its neighbour. The plot width divided by this gap gives the
# label count, so it scales with -Width rather than a hardcoded count that
# would crowd a narrower chart or under-fill a wider one. Chosen so the
# default 1080px chart (plot width ~1012px after padding) lands around 7
# labels, inside the 6-8 range asked for.
$script:AreaChartLabelSpacingPx = 150

function ConvertTo-HtmlEncoded {
    <#
    .SYNOPSIS
        Escapes text for safe inclusion in HTML or SVG.
    .DESCRIPTION
        Delivery group and application names come from the customer's live
        site and routinely contain ampersands, and could contain angle
        brackets. Unescaped, they corrupt the document or, worse, inject
        markup into a report that is opened in a browser.

        Unicode bidirectional-override and isolate control characters
        (U+202A-U+202E, U+2066-U+2069) are stripped rather than escaped.
        They carry no legitimate use in a plain label, and left in place
        they can visually reorder the rendered text without changing its
        underlying characters - in a report a customer reads and makes a
        licensing decision from, a crafted delivery-group name using one
        could make the page display something other than what it contains.

        A label that is empty, or becomes whitespace-only once any bidi
        controls are stripped, is normalised to the empty string: it carries
        no visible identity, and callers treat an empty result the same as
        an absent label (e.g. omitting a chart's <title>) rather than
        rendering an element with nothing readable in it.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([AllowNull()][string] $Text)

    if ([string]::IsNullOrEmpty($Text)) { return '' }

    $clean = [regex]::Replace($Text, '[\u202A-\u202E\u2066-\u2069]', '')
    if ([string]::IsNullOrWhiteSpace($clean)) { return '' }

    return $clean.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;').
        Replace('"', '&quot;').Replace("'", '&#39;')
}

function Get-TruncatedLabel {
    <#
    .SYNOPSIS
        Shortens a label to fit an approximate pixel width, with an ellipsis.
    .DESCRIPTION
        SVG has no text-measurement API available to us, so width is
        estimated from the font size: roughly 0.55 * FontSizePx per
        character is a reasonable average for these families. Erring narrow
        is safer than overprinting - an ellipsis is recoverable, overprinted
        text is not, and the full name is always present in the adjacent
        table in the report.

        Truncation must happen on the raw label, before ConvertTo-HtmlEncoded
        runs. Truncating an already-escaped string risks cutting an entity
        like "&amp;" in half, which both breaks markup and mis-estimates the
        rendered width (an entity is several source characters wide but one
        rendered glyph).
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string] $Text,
        [Parameter(Mandatory)][double] $MaxWidthPx,
        [double] $FontSizePx = 13
    )

    if ([string]::IsNullOrEmpty($Text)) { return $Text }

    $ellipsis = [string] [char] 0x2026
    $charW = 0.55 * $FontSizePx
    if ($charW -le 0) { return $Text }

    $maxChars = [math]::Floor($MaxWidthPx / $charW)
    if ($Text.Length -le $maxChars) { return $Text }
    if ($maxChars -le 1) { return $ellipsis }

    return $Text.Substring(0, [int] ($maxChars - 1)) + $ellipsis
}

function Get-ChartSeriesColor {
    <#
    .SYNOPSIS
        Colour for one chart series, drawn from the mark-safe subset of the
        brand palette.
    .DESCRIPTION
        Only four of the twelve brand tokens are legible as a data mark on
        the white card the charts sit on: Cobalt 50, Jasper 80, Cobalt 70
        and Jasper 100 all clear 3:1 contrast against white. Mint 40,
        Slate 20 and Aether 20 do not (as low as 1.21:1) and are never
        returned here, even though they keep their brand-sanctioned roles
        as background and accent tints elsewhere.

        The ramp is fixed order, not cycled: index 0 is always Cobalt 50,
        index 1 always Jasper 80, and so on. An index past the end of the
        ramp does NOT wrap around to reuse an earlier hue - cycling a
        palette by index is a data-visualisation anti-pattern in its own
        right, and here it would silently give two unrelated series the
        same colour. Instead the last mark-safe hue is returned, and the
        caller is expected to have grouped any remainder into an "Other"
        bucket before it gets this far. Every caller in this project caps
        its series count at or below the ramp length already, so this is a
        guard rail rather than a path any current caller takes.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)][int] $Index,
        [int] $SeriesCount = 1
    )

    $p = Get-BrandPalette
    $ramp = @($p['Cobalt50'], $p['Jasper80'], $p['Cobalt70'], $p['Jasper100'])

    if ($Index -lt 0) { return $ramp[0] }
    if ($Index -ge $ramp.Count) { return $ramp[$ramp.Count - 1] }
    return $ramp[$Index]
}

function New-SvgAreaChart {
    <#
    .SYNOPSIS
        Filled area chart of a concurrency time series.
    .DESCRIPTION
        Single-series by nature (one line of concurrency over time), so it
        uses exactly one hue: Cobalt 50 at low opacity for the fill, full
        strength for the stroke.

        When -Label is supplied, it is rendered as an SVG <title> element -
        the first child of the <svg>, which is better supported by screen
        readers than an aria-label attribute on inline SVG. When no label is
        supplied (or it is empty/whitespace-only), the <title> element and
        any aria attribute are omitted entirely rather than emitting one
        that declares an empty accessible name.

        X-axis labels: a reader who can only see the first and last date on
        a spike-and-drop chart can tell something happened but not when, so
        intermediate labels are placed across the axis rather than only at
        the two ends. The label count is derived from the plot width, not
        hardcoded - $script:AreaChartLabelSpacingPx is the minimum pixel
        gap a label is allowed, and the plot width divided by that gap gives
        the count. At the default 1080px width that lands in the 6-8 label
        range the owner asked for; a narrower chart gets fewer, a wider one
        more, and neither ever crowds. Label positions are then chosen as
        evenly spaced INDICES into the actual data array (deduplicated), so
        a label is always drawn under a real data point's x coordinate
        rather than an interpolated one - this is also what keeps a
        two-or-three-point series from emitting duplicate labels: the
        desired count is capped at the point count and rounding collisions
        are collapsed by -Unique.

        The label date format is chosen from the span the points actually
        cover, not assumed: a series spanning a day or less (the
        15-minute-bucket concurrency series over a short window) uses
        "HH:mm", because "dd MMM" would print the same date under every
        label. A series spanning more than a day (the daily-trend chart, or
        concurrency over a multi-day window) uses "dd MMM".

        Vertical gridlines are drawn at the same label positions, styled
        identically to the existing horizontal gridlines (Slate 20 at 0.45
        opacity) so they read as structure, not data.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()] $Points,
        [int] $Width = 1080,
        [int] $Height = 260,
        [string] $Label = ''
    )

    $p = Get-BrandPalette
    $padL = 52; $padR = 16; $padT = 16; $padB = 30
    $plotW = $Width - $padL - $padR
    $plotH = $Height - $padT - $padB

    $titleText = ConvertTo-HtmlEncoded -Text $Label
    $titleEl = if ($titleText) { "  <title>$titleText</title>`n" } else { '' }

    $data = @($Points)
    if ($data.Count -eq 0) {
        return @"
<svg viewBox="0 0 $Width $Height" width="100%" xmlns="http://www.w3.org/2000/svg" role="img">
$titleEl  <text x="$($Width/2)" y="$($Height/2)" text-anchor="middle" font-size="14" fill="$($p['Jasper80'])">No session activity in this window</text>
</svg>
"@
    }

    $max = ($data | Measure-Object -Property Count -Maximum).Maximum
    # A flat series makes the range zero. Without this floor the scale divides
    # by zero and every coordinate becomes NaN, rendering a blank chart.
    if ($max -le 0) { $max = 1 }

    $stepX = if ($data.Count -gt 1) { $plotW / ($data.Count - 1) } else { 0 }

    $coords = for ($i = 0; $i -lt $data.Count; $i++) {
        $x = $padL + ($i * $stepX)
        $y = $padT + $plotH - (($data[$i].Count / $max) * $plotH)
        '{0:F1},{1:F1}' -f $x, $y
    }

    $polyline = $coords -join ' '
    $polygon = "$padL,$($padT + $plotH) $polyline $($padL + $plotW),$($padT + $plotH)"

    # Four gridlines is enough to read a value without competing with the data.
    $gridlines = for ($g = 0; $g -le 4; $g++) {
        $y = $padT + $plotH - (($g / 4.0) * $plotH)
        $v = [int][math]::Round(($g / 4.0) * $max)
        @"
  <line x1="$padL" y1="$([math]::Round($y,1))" x2="$($padL + $plotW)" y2="$([math]::Round($y,1))" stroke="$($p['Slate20'])" stroke-width="1" opacity="0.45"/>
  <text x="$($padL - 8)" y="$([math]::Round($y + 4,1))" text-anchor="end" font-size="11" fill="$($p['Jasper80'])">$v</text>
"@
    }

    # How many x-axis labels fit without crowding, derived from the plot
    # width rather than hardcoded (see $script:AreaChartLabelSpacingPx).
    # Never more than one label per data point.
    $desiredLabelCount = [math]::Max(2, [int][math]::Floor($plotW / $script:AreaChartLabelSpacingPx) + 1)
    $labelCount = [math]::Min($desiredLabelCount, $data.Count)
    if ($labelCount -lt 1) { $labelCount = 1 }

    # Evenly spaced INDICES into the real data array, so every label sits
    # under an actual data point's x coordinate. -Unique collapses any
    # rounding collisions, which is what keeps a very-short series (2-3
    # points asked to carry more labels than it has room for) from ever
    # emitting a duplicate.
    $labelIndices = if ($labelCount -le 1) {
        @(0)
    } else {
        @(0..($labelCount - 1) | ForEach-Object {
            [int][math]::Round($_ * ($data.Count - 1) / [double]($labelCount - 1))
        } | Select-Object -Unique)
    }

    # A span of a day or less (a short concurrency window sampled every 15
    # minutes) needs a time, not a date - "dd MMM" would print the same
    # calendar date under every label. Anything longer (the daily-trend
    # chart, or concurrency over a multi-day window) uses a date.
    $spanHours = if ($data.Count -gt 1) { ($data[-1].TimeUtc - $data[0].TimeUtc).TotalHours } else { 0 }
    $axisDateFormat = if ($spanHours -gt 0 -and $spanHours -le 24) { 'HH:mm' } else { 'dd MMM' }

    $xAxis = for ($li = 0; $li -lt $labelIndices.Count; $li++) {
        $idx = $labelIndices[$li]
        $x = $padL + ($idx * $stepX)
        $anchor = if ($li -eq 0) { 'start' } elseif ($li -eq $labelIndices.Count - 1) { 'end' } else { 'middle' }
        $labelText = ConvertTo-HtmlEncoded -Text $data[$idx].TimeUtc.ToString($axisDateFormat)
        @"
  <line x1="$([math]::Round($x,1))" y1="$padT" x2="$([math]::Round($x,1))" y2="$($padT + $plotH)" stroke="$($p['Slate20'])" stroke-width="1" opacity="0.45"/>
  <text x="$([math]::Round($x,1))" y="$($Height - 8)" text-anchor="$anchor" font-size="11" fill="$($p['Jasper80'])">$labelText</text>
"@
    }

    @"
<svg viewBox="0 0 $Width $Height" width="100%" xmlns="http://www.w3.org/2000/svg" role="img">
$titleEl$($gridlines -join "`n")
$($xAxis -join "`n")
  <polygon points="$polygon" fill="$($p['Cobalt50'])" opacity="0.16"/>
  <polyline points="$polyline" fill="none" stroke="$($p['Cobalt50'])" stroke-width="2" stroke-linejoin="round"/>
</svg>
"@
}

function New-SvgBarChart {
    <#
    .SYNOPSIS
        Horizontal bar chart with value labels.
    .DESCRIPTION
        Horizontal rather than vertical, because delivery group and
        application names are long and would otherwise need rotated,
        unreadable labels. Because those names are long, each label is
        truncated (with a trailing ellipsis) to fit its column via
        Get-TruncatedLabel before being escaped - the label is drawn before
        the bar in document order, so an untruncated long name would be
        painted over by the bar's opaque fill with no indication anything
        was cut off. The full name is always available in the report's
        adjacent data table.

        Every bar is drawn in the same colour, Cobalt 50. Bars in this chart
        show magnitude, not identity - the label beside each bar already
        carries identity, so colouring bars individually would add nothing
        and would break the brand's four-hue-per-composition rule the moment
        a delivery group list runs past a handful of rows.

        When -Label is supplied, it is rendered as an SVG <title> element,
        the chart's accessible name for screen readers. It is omitted
        entirely (not emitted empty) when no label is given.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()] $Items,
        [int] $Width = 1080,
        [int] $Height = 0,
        [string] $ValueProperty = 'Value',
        [string] $LabelProperty = 'Label',
        [string] $Label = ''
    )

    $p = Get-BrandPalette
    $data = @($Items)

    $titleText = ConvertTo-HtmlEncoded -Text $Label
    $titleEl = if ($titleText) { "  <title>$titleText</title>`n" } else { '' }

    if ($data.Count -eq 0) {
        return @"
<svg viewBox="0 0 $Width 60" width="100%" xmlns="http://www.w3.org/2000/svg" role="img">
$titleEl  <text x="$($Width/2)" y="34" text-anchor="middle" font-size="14" fill="$($p['Jasper80'])">No data available</text>
</svg>
"@
    }

    $rowH = 34
    $gap = 8
    if ($Height -le 0) { $Height = ($data.Count * ($rowH + $gap)) + 16 }

    $labelW = 260
    $valueW = 70
    $barMaxW = $Width - $labelW - $valueW - 24
    $labelFontSize = 13
    $labelMaxWidthPx = $labelW - 10

    $max = ($data | ForEach-Object { [double] $_.$ValueProperty } | Measure-Object -Maximum).Maximum
    if ($max -le 0) { $max = 1 }

    # Every bar carries the same colour deliberately - see the function's
    # .DESCRIPTION. Do not colour bars by index.
    $barColor = $p['Cobalt50']

    $rows = for ($i = 0; $i -lt $data.Count; $i++) {
        $value = [double] $data[$i].$ValueProperty
        $rawLabel = [string] $data[$i].$LabelProperty
        $truncated = Get-TruncatedLabel -Text $rawLabel -MaxWidthPx $labelMaxWidthPx -FontSizePx $labelFontSize
        $label = ConvertTo-HtmlEncoded -Text $truncated
        $y = 8 + ($i * ($rowH + $gap))
        $barW = [math]::Max(2, ($value / $max) * $barMaxW)

        @"
  <text x="0" y="$($y + 21)" font-size="$labelFontSize" fill="$($p['Jasper100'])">$label</text>
  <rect x="$labelW" y="$y" width="$([math]::Round($barW,1))" height="$rowH" rx="4" fill="$barColor"/>
  <text x="$($labelW + $barW + 10)" y="$($y + 21)" font-size="13" font-weight="600" fill="$($p['Jasper100'])">$([int]$value)</text>
"@
    }

    @"
<svg viewBox="0 0 $Width $Height" width="100%" xmlns="http://www.w3.org/2000/svg" role="img">
$titleEl$($rows -join "`n")
</svg>
"@
}

function New-SvgGroupedBarChart {
    <#
    .SYNOPSIS
        Two-series grouped bar chart, for comparing windows side by side.
    .DESCRIPTION
        Limited to two series deliberately: Cobalt 50 and Jasper 80. That
        pair was validated for CVD separation (Delta E 21.6 deutan, 8.9
        tritan, 22.7 normal vision) and both clear 3:1 contrast against the
        white card, which is what makes a two-hue comparison safe for
        colour-blind readers here - together with the per-bar value labels
        and legend this chart keeps as the secondary encoding. It also keeps
        the composition inside the brand's four-hue-per-composition rule.

        Both the x-axis group labels and the legend entries are truncated
        via Get-TruncatedLabel before being escaped, for the same reason as
        the bar chart: names are long, and at six to ten groups the per-
        group column and the legend's available width both run well short
        of a realistic name. The legend lays itself out left to right from a
        running total of each item's estimated rendered width (swatch plus
        text) rather than a fixed pitch, so two long series names can never
        overlap regardless of how long they are.

        When -Label is supplied, it is rendered as an SVG <title> element,
        the chart's accessible name for screen readers. It is omitted
        entirely (not emitted empty) when no label is given.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()] $Items,
        [Parameter(Mandatory)][string[]] $Series,
        [int] $Width = 1080,
        [int] $Height = 300,
        [string] $Label = ''
    )

    $p = Get-BrandPalette
    $data = @($Items)
    if ($data.Count -eq 0 -or $Series.Count -eq 0) {
        return New-SvgBarChart -Items @() -Width $Width -Label $Label
    }

    $titleText = ConvertTo-HtmlEncoded -Text $Label
    $titleEl = if ($titleText) { "  <title>$titleText</title>`n" } else { '' }

    $padL = 52; $padB = 46; $padT = 16; $padR = 16
    $plotW = $Width - $padL - $padR
    $plotH = $Height - $padT - $padB

    $allValues = foreach ($item in $data) { foreach ($s in $Series) { [double] $item.$s } }
    $max = ($allValues | Measure-Object -Maximum).Maximum
    if ($max -le 0) { $max = 1 }

    $groupW = $plotW / $data.Count
    $barW = ($groupW * 0.7) / $Series.Count
    $axisLabelFontSize = 12
    $axisLabelMaxWidthPx = $groupW - 8

    $bars = for ($g = 0; $g -lt $data.Count; $g++) {
        $groupX = $padL + ($g * $groupW) + ($groupW * 0.15)
        $rawLabel = [string] $data[$g].Label
        $truncatedLabel = Get-TruncatedLabel -Text $rawLabel -MaxWidthPx $axisLabelMaxWidthPx -FontSizePx $axisLabelFontSize
        $label = ConvertTo-HtmlEncoded -Text $truncatedLabel

        $inner = for ($s = 0; $s -lt $Series.Count; $s++) {
            $value = [double] $data[$g].($Series[$s])
            $h = ($value / $max) * $plotH
            $x = $groupX + ($s * $barW)
            $y = $padT + $plotH - $h
            $color = Get-ChartSeriesColor -Index $s -SeriesCount $Series.Count
            @"
  <rect x="$([math]::Round($x,1))" y="$([math]::Round($y,1))" width="$([math]::Round($barW - 3,1))" height="$([math]::Round($h,1))" rx="3" fill="$color"/>
  <text x="$([math]::Round($x + ($barW/2) - 2,1))" y="$([math]::Round($y - 5,1))" text-anchor="middle" font-size="11" font-weight="600" fill="$($p['Jasper100'])">$([int]$value)</text>
"@
        }

        @"
$($inner -join "`n")
  <text x="$([math]::Round($groupX + ($groupW * 0.35),1))" y="$($padT + $plotH + 20)" text-anchor="middle" font-size="$axisLabelFontSize" fill="$($p['Jasper100'])">$label</text>
"@
    }

    # Legend items are laid out left to right from a running total of each
    # item's estimated rendered width (swatch + gap + text + trailing gap),
    # not a fixed pitch - a fixed pitch overlaps as soon as a series name
    # runs longer than the pitch allows.
    $legendFontSize = 12
    $legendMaxLabelWidthPx = 160
    $legendX = $padL

    $legend = for ($s = 0; $s -lt $Series.Count; $s++) {
        $color = Get-ChartSeriesColor -Index $s -SeriesCount $Series.Count
        $rawSeriesLabel = [string] $Series[$s]
        $truncatedSeriesLabel = Get-TruncatedLabel -Text $rawSeriesLabel -MaxWidthPx $legendMaxLabelWidthPx -FontSizePx $legendFontSize
        $seriesLabel = ConvertTo-HtmlEncoded -Text $truncatedSeriesLabel

        $x = $legendX
        @"
  <rect x="$x" y="$($Height - 16)" width="12" height="12" rx="2" fill="$color"/>
  <text x="$($x + 18)" y="$($Height - 6)" font-size="$legendFontSize" fill="$($p['Jasper80'])">$seriesLabel</text>
"@
        $textWidthEstimate = $truncatedSeriesLabel.Length * (0.55 * $legendFontSize)
        $legendX += 18 + $textWidthEstimate + 24
    }

    @"
<svg viewBox="0 0 $Width $Height" width="100%" xmlns="http://www.w3.org/2000/svg" role="img">
$titleEl  <line x1="$padL" y1="$($padT + $plotH)" x2="$($padL + $plotW)" y2="$($padT + $plotH)" stroke="$($p['Slate20'])" stroke-width="1"/>
$($bars -join "`n")
$($legend -join "`n")
</svg>
"@
}

# endregion 70-Charts.ps1

# ----------------------------------------------------------------------------
# region 75-HtmlReport.ps1
# ----------------------------------------------------------------------------
# ============================================================================
#  HTML report
#
#  Assembles a single self-contained document. Every asset is inline, so the
#  report opens on an air-gapped machine, survives being emailed, and prints.
# ============================================================================

function Format-LocalTime {
    <#
    .SYNOPSIS
        Renders a UTC timestamp in a named timezone, with the zone shown.
    .DESCRIPTION
        Every displayed time names its timezone. A licensing figure attached to
        an unqualified timestamp invites the reader to assume their own zone,
        and peak times shift by hours between zones.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [AllowNull()][Nullable[datetime]] $Utc,
        [string] $TimeZoneId = ([System.TimeZoneInfo]::Local.Id)
    )

    if (-not $Utc) { return 'n/a' }

    try {
        $tz = [System.TimeZoneInfo]::FindSystemTimeZoneById($TimeZoneId)
        $local = [System.TimeZoneInfo]::ConvertTimeFromUtc($Utc, $tz)
        $abbr = if ($tz.IsDaylightSavingTime($local)) { $tz.DaylightName } else { $tz.StandardName }
        return '{0:dd MMM yyyy HH:mm} ({1})' -f $local, $abbr
    } catch {
        # An unknown timezone id must not sink the whole report.
        return '{0:dd MMM yyyy HH:mm} (UTC)' -f $Utc
    }
}

function New-TruncationNotice {
    <#
    .SYNOPSIS
        The banner shown against any window wider than the retained history.
    .DESCRIPTION
        Rendered next to the figures it affects rather than as a footnote,
        because an under-counted window is invisible in the output otherwise
        and would understate the customer's licence need.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Mandatory)][pscustomobject] $Window)

    if (-not $Window.IsTruncated) { return '' }

    @"
<div class="truncated">
  <strong>This $($Window.Days)-day window is truncated.</strong>
  Only $($Window.AvailableDays) day(s) of session history are available on this
  site, so the figures below cover $($Window.AvailableDays) days rather than
  $($Window.Days). Treat them as a <strong>lower bound</strong>, not a true
  count: the real $($Window.Days)-day unique user total is higher than shown.
  Citrix grooms raw session data at 90 days on Premium, 31 days on Advanced,
  and 7 days on other editions; history can also be shorter simply because the
  site had no activity further back.
</div>
"@
}

function New-RetentionUncertainNotice {
    <#
    .SYNOPSIS
        The banner shown when retention itself could not be confirmed.
    .DESCRIPTION
        Invoke-AuditPreflight normally measures retention from the oldest
        ENDED session, because grooming only ever removes ended sessions.
        When a site has zero ended sessions within retention -- a handful of
        always-connected persistent desktops on a small, non-Premium site is
        a real, unremarkable example, not an edge case -- it falls back to
        the oldest session START date instead, which can overstate true
        retention if older sessions have already been groomed away.

        In that fallback, every requested window can nominally fit inside
        the reported AvailableDays with none of them individually flagged as
        truncated, so New-TruncationNotice alone renders nothing and the
        report would otherwise show unqualified figures next to an
        unqualified "Session history retained" line -- an affirmative
        coverage claim that was never actually established. This banner
        exists to carry that caveat into the report itself, the same way
        New-TruncationNotice carries a confirmed shortfall into it, because a
        warning that only reaches audit.log never reaches the customer who
        acts on the report.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([AllowNull()][pscustomobject] $Preflight)

    if (-not $Preflight) { return '' }
    if ($Preflight.RetentionBasis -eq 'OldestEndedSession') { return '' }

    $detail = if ($Preflight.RetentionBasis -eq 'NoSessions') {
        'No sessions of any kind were returned by the history probe, so how much session history this site retains is entirely unknown.'
    } else {
        "No completed (ended) session was found within the probed history, so retention could not be confirmed from Citrix's own grooming signal. The $($Preflight.AvailableDays)-day figure shown is measured from the oldest still-open session instead, which can overstate true retention if older sessions have already been groomed away."
    }

    @"
<div class="truncated">
  <strong>Session history retention could not be confirmed.</strong>
  $detail
  Every window requested for this report nominally fits inside that figure,
  so no individual window below is marked truncated -- but treat all figures
  in this report as a <strong>lower bound</strong>, not a confirmed true
  count, until retention is verified directly (for example against
  Set-MonitorConfiguration on this site).
</div>
"@
}

function New-FetchWarningNotice {
    <#
    .SYNOPSIS
        The banner shown when an entity's data was fetched incompletely.
    .DESCRIPTION
        An arbitrarily truncated fetch produces an under-count that looks
        exactly like a complete figure. Before this existed the only trace was
        a line in audit.log, which nobody reading the report ever sees -- the
        same silent-undercount failure the retention banner exists to prevent,
        applied to Connections and ApplicationInstances instead of Sessions.

        Rendered inside the section it affects (client devices, applications)
        and again at the top of the report, so it cannot be missed whichever
        figure the reader goes to.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][AllowNull()] $Warnings,
        [string[]] $Entities
    )

    $relevant = @($Warnings | Where-Object {
        $_ -and (-not $Entities -or $Entities -contains $_.Entity)
    })
    if ($relevant.Count -eq 0) { return '' }

    $items = ($relevant | ForEach-Object {
        "<li>$(ConvertTo-HtmlEncoded -Text $_.Message)</li>"
    }) -join "`n"

    @"
<div class="truncated">
  <strong>Some data for this report was not fetched completely.</strong>
  <ul>
$items
  </ul>
  Figures derived from the affected data are a <strong>lower bound</strong>,
  not a true count. Re-running with a narrower reporting window usually
  resolves it.
</div>
"@
}

function New-WindowSection {
    <#
    .SYNOPSIS
        One reporting window rendered as a full section.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)][pscustomobject] $Window,
        [Parameter(Mandatory)][string] $TimeZoneId,
        [AllowEmptyCollection()][AllowNull()] $FetchWarnings
    )

    $c = $Window.Concurrency

    # What "business hours" actually means, stated wherever a business-hours
    # figure appears. Without it a reader in another zone assumes their own,
    # and the same number means different things at 08:00-18:00 London and
    # 08:00-18:00 Sydney.
    $bizStart = if ($null -ne $c.BusinessHourStart) { [int] $c.BusinessHourStart } else { 8 }
    $bizEnd   = if ($null -ne $c.BusinessHourEnd)   { [int] $c.BusinessHourEnd }   else { 18 }
    $bizZone  = if ($c.BusinessTimeZoneId) { [string] $c.BusinessTimeZoneId } else { 'UTC' }
    $bizLabel = '{0:00}:00&ndash;{1:00}:00 Mon&ndash;Fri, {2}' -f $bizStart, $bizEnd,
        (ConvertTo-HtmlEncoded -Text $bizZone)

    $sections = New-Object System.Collections.Generic.List[string]

    [void] $sections.Add(@"
<div class="card">
  <h2>$($Window.Days)-day window</h2>
  <p class="footnote">
    $(Format-LocalTime -Utc $Window.WindowStartUtc -TimeZoneId $TimeZoneId)
    &rarr; $(Format-LocalTime -Utc $Window.WindowEndUtc -TimeZoneId $TimeZoneId)
  </p>
  $(New-TruncationNotice -Window $Window)

  <div class="kpi-row">
    <div class="kpi">
      <div class="label">Unique users</div>
      <div class="value">$($Window.UniqueUsers)</div>
      <div class="note">Distinct users with at least one session</div>
    </div>
    <div class="kpi">
      <div class="label">Peak concurrent</div>
      <div class="value">$($c.Peak)</div>
      <div class="note">$(Format-LocalTime -Utc $c.PeakAtUtc -TimeZoneId $TimeZoneId)</div>
    </div>
    <div class="kpi">
      <div class="label">95th percentile concurrent</div>
      <div class="value">$([math]::Round($c.P95, 0))</div>
      <div class="note">Business hours ($bizLabel): $([math]::Round($c.BusinessP95, 0))</div>
    </div>
    <div class="kpi">
      <div class="label">Total sessions</div>
      <div class="value">$($Window.TotalSessions)</div>
      <div class="note">$($Window.AnonymousSessions) anonymous, $($Window.UnattributedSessions) unattributed</div>
    </div>
  </div>

  <h3>Concurrent sessions over time</h3>
  $(New-SvgAreaChart -Points $c.Series -Width 1080 -Height 260 -Label "Concurrent sessions over $($Window.Days) days")

  <h3>How often each concurrency level was reached</h3>
  <p class="footnote">
    This table tells a brief spike apart from sustained load: each row
    states how often concurrent sessions reached or stayed at or below a
    given level, from typical use up to the single highest moment recorded.
    A peak that sits far above the rest of the distribution usually reflects
    one unusual moment, not a level the site ran at routinely.
  </p>
  <div class="scroll-x">
  <table>
    <thead><tr><th>Concurrent sessions were&hellip;</th><th class="num">All hours</th><th class="num">Business hours<br><span class="footnote">$bizLabel</span></th></tr></thead>
    <tbody>
      <tr><td>Half the time, at or below</td><td class="num">$([math]::Round($c.P50,0))</td><td class="num">&mdash;</td></tr>
      <tr><td>9 times out of 10, at or below</td><td class="num">$([math]::Round($c.P90,0))</td><td class="num">&mdash;</td></tr>
      <tr><td>19 times out of 20, at or below</td><td class="num">$([math]::Round($c.P95,0))</td><td class="num">$([math]::Round($c.BusinessP95,0))</td></tr>
      <tr><td>99 times out of 100, at or below</td><td class="num">$([math]::Round($c.P99,0))</td><td class="num">&mdash;</td></tr>
      <tr><td>Highest at any single moment</td><td class="num">$($c.Peak)</td><td class="num">$($c.BusinessPeak)</td></tr>
    </tbody>
  </table>
  </div>
</div>
"@)

    if ($Window.DeliveryGroups -and $Window.DeliveryGroups.Count -gt 0) {
        $rows = ($Window.DeliveryGroups | ForEach-Object {
            "<tr><td>$(ConvertTo-HtmlEncoded -Text $_.Name)</td><td class=""num"">$($_.UniqueUsers)</td><td class=""num"">$($_.PeakConcurrent)</td><td class=""num"">$($_.TotalSessions)</td></tr>"
        }) -join "`n"

        $chartItems = @($Window.DeliveryGroups | Select-Object -First 10 | ForEach-Object {
            [pscustomobject]@{ Label = $_.Name; Value = $_.UniqueUsers }
        })

        [void] $sections.Add(@"
<div class="card">
  <h2>By delivery group &mdash; $($Window.Days) days</h2>
  $(New-SvgBarChart -Items $chartItems -Width 1080 -Label "Unique users by delivery group over $($Window.Days) days")
  <div class="scroll-x">
  <table>
    <thead><tr><th>Delivery group</th><th class="num">Unique users</th><th class="num">Peak concurrent</th><th class="num">Sessions</th></tr></thead>
    <tbody>$rows</tbody>
  </table>
  </div>
</div>
"@)
    }

    if ($Window.Applications -and $Window.Applications.Count -gt 0) {
        $st = $Window.SessionTypes
        $appRows = ($Window.Applications | ForEach-Object {
            "<tr><td>$(ConvertTo-HtmlEncoded -Text $_.Name)</td><td class=""num"">$($_.UniqueUsers)</td><td class=""num"">$($_.Launches)</td></tr>"
        }) -join "`n"

        [void] $sections.Add(@"
<div class="card">
  <h2>Applications and session types &mdash; $($Window.Days) days</h2>
  $(New-FetchWarningNotice -Warnings $FetchWarnings -Entities @('ApplicationInstances','Applications'))
  <div class="kpi-row">
    <div class="kpi"><div class="label">Desktop users</div><div class="value">$($st.DesktopUsers)</div><div class="note">$($st.DesktopSessions) sessions</div></div>
    <div class="kpi"><div class="label">Application users</div><div class="value">$($st.ApplicationUsers)</div><div class="note">$($st.ApplicationSessions) sessions</div></div>
  </div>
  <h3>Top published applications</h3>
  <div class="scroll-x">
  <table>
    <thead><tr><th>Application</th><th class="num">Unique users</th><th class="num">Launches</th></tr></thead>
    <tbody>$appRows</tbody>
  </table>
  </div>
</div>
"@)
    }

    if ($Window.ClientDevices -and $Window.ClientDevices.UniqueDevices -gt 0) {
        $cd = $Window.ClientDevices
        $verRows = ($cd.VersionCounts | Select-Object -First 12 | ForEach-Object {
            "<tr><td>$(ConvertTo-HtmlEncoded -Text $_.Version)</td><td class=""num"">$($_.Count)</td></tr>"
        }) -join "`n"

        [void] $sections.Add(@"
<div class="card">
  <h2>Client devices &mdash; $($Window.Days) days</h2>
  $(New-FetchWarningNotice -Warnings $FetchWarnings -Entities @('Connections'))
  <div class="kpi-row">
    <div class="kpi"><div class="label">Unique devices</div><div class="value">$($cd.UniqueDevices)</div><div class="note">By client name</div></div>
    <div class="kpi"><div class="label">Unique addresses</div><div class="value">$($cd.UniqueAddresses)</div><div class="note">By client IP</div></div>
  </div>
  <h3>Citrix Workspace app versions</h3>
  <div class="scroll-x">
  <table>
    <thead><tr><th>Version</th><th class="num">Connections</th></tr></thead>
    <tbody>$verRows</tbody>
  </table>
  </div>
</div>
"@)
    }

    if ($Window.DailyTrend -and $Window.DailyTrend.Count -gt 0) {
        $trendUsers = @($Window.DailyTrend | ForEach-Object {
            [pscustomobject]@{ TimeUtc = $_.DateUtc; Count = $_.UniqueUsers }
        })
        [void] $sections.Add(@"
<div class="card">
  <h2>Daily activity &mdash; $($Window.Days) days</h2>
  <h3>Unique users per day</h3>
  $(New-SvgAreaChart -Points $trendUsers -Width 1080 -Height 220 -Label "Unique users per day over $($Window.Days) days")
</div>
"@)
    }

    return ($sections -join "`n")
}

function New-HtmlReport {
    <#
    .SYNOPSIS
        Builds the complete self-contained HTML report.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Mandatory)][pscustomobject] $Analysis)

    $cfg = $Analysis.Config
    $tz = $cfg.DisplayTimeZoneId
    $widest = $Analysis.Windows | Sort-Object Days -Descending | Select-Object -First 1

    $demoBanner = if ($Analysis.IsDemo) {
        @"
<div class="notice-demo">
  <strong>This is a synthetic demonstration report.</strong>
  The figures below were generated from randomly created sample data and
  describe no real Citrix environment. Do not use them for any licensing
  decision.
</div>
"@
    } else { '' }

    # Carried on Config (see 20-Config.ps1) rather than plumbed through the
    # fetch-warning list: this is not a fetch that came back incomplete, it
    # is a plain fact about how the whole run was transported, so it belongs
    # in its own banner with its own accurate wording -- reusing the generic
    # "some data was not fetched completely" copy here would misdescribe the
    # risk. Rendered wherever the report is opened, not only in audit.log,
    # because a customer forwarding the report to a consultant needs to see
    # this without also having the run log in hand.
    $insecureBanner = if ($cfg.PSObject.Properties['IsHttp'] -and $cfg.IsHttp) {
        $guidanceUrl = ConvertTo-HtmlEncoded -Text $script:MonitorTlsGuidanceUrl
        @"
<div class="truncated">
  <strong>This report was generated over an unencrypted HTTP connection.</strong>
  See Citrix's guidance on securing the Monitor Service with TLS:
  <a href="$guidanceUrl">$guidanceUrl</a>
</div>
"@
    } else { '' }

    $fetchWarnings = @()
    if ($Analysis.PSObject.Properties['FetchWarnings']) { $fetchWarnings = @($Analysis.FetchWarnings) }
    $fetchBanner = New-FetchWarningNotice -Warnings $fetchWarnings
    $retentionBanner = New-RetentionUncertainNotice -Preflight $Analysis.Preflight

    $retentionCell = if ($Analysis.Preflight) {
        $days = "$($Analysis.Preflight.AvailableDays) days"
        if ($Analysis.Preflight.RetentionBasis -ne 'OldestEndedSession') {
            "$days (unconfirmed &mdash; lower bound)"
        } else {
            $days
        }
    } else { 'unknown' }

    $windowSections = ($Analysis.Windows | Sort-Object Days | ForEach-Object {
        New-WindowSection -Window $_ -TimeZoneId $tz -FetchWarnings $fetchWarnings
    }) -join "`n"

    $comparison = @($Analysis.Windows | Sort-Object Days | ForEach-Object {
        [pscustomobject]@{
            Label = "$($_.Days) days"
            'Unique users' = $_.UniqueUsers
            'Peak concurrent' = $_.Concurrency.Peak
        }
    })

    $logo = Get-BrandLogoSvg -Variant white
    $css = Get-BrandCss

    @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Citrix Usage Report</title>
<style>
$css
</style>
</head>
<body>

<header class="report-header">
  <div class="wrap header-row">
    <div>
      <div class="logo">$logo</div>
      <h1>Citrix usage report</h1>
      <p class="sub">Unique users and concurrent sessions over the requested periods</p>
    </div>
    <div class="header-actions no-print">
      <button type="button" id="print-btn" class="print-btn">Save as PDF</button>
      <p class="print-hint">or press Ctrl+P</p>
    </div>
  </div>
</header>

<div class="wrap">

$demoBanner

$insecureBanner

$fetchBanner

$retentionBanner

<div class="card">
  <h2>Summary</h2>
  <div class="scroll-x">
  <table>
    <tbody>
      <tr><td>Environment</td><td>$(ConvertTo-HtmlEncoded -Text $cfg.EnvironmentLabel)</td></tr>
      <tr><td>Report generated</td><td>$(Format-LocalTime -Utc $Analysis.GeneratedUtc -TimeZoneId $tz)</td></tr>
      <tr><td>Session history retained</td><td>$retentionCell</td></tr>
      <tr><td>Delivery groups</td><td>$($Analysis.SiteTotals.DeliveryGroups)</td></tr>
      <tr><td>Machines</td><td>$($Analysis.SiteTotals.Machines)</td></tr>
    </tbody>
  </table>
  </div>

  <h3>Headline, based on the $($widest.Days)-day window</h3>
  $(New-TruncationNotice -Window $widest)
  <div class="kpi-row">
    <div class="kpi">
      <div class="label">Unique users</div>
      <div class="value">$($widest.UniqueUsers)</div>
      <div class="note">Distinct users with at least one session</div>
    </div>
    <div class="kpi">
      <div class="label">Peak concurrent</div>
      <div class="value">$($widest.Concurrency.Peak)</div>
      <div class="note">$(Format-LocalTime -Utc $widest.Concurrency.PeakAtUtc -TimeZoneId $tz)</div>
    </div>
  </div>

  <h3>Windows compared</h3>
  $(New-SvgGroupedBarChart -Items $comparison -Series @('Unique users','Peak concurrent') -Width 1080 -Height 300 -Label 'Unique users and peak concurrency compared across windows')
</div>

$windowSections

<div class="card">
  <h2>How to read this report</h2>
  <p><strong>Unique users</strong> counts distinct users with at least one
  session overlapping the window. A session that began before the window still
  counts, because it still occupies a licence inside it.</p>
  <p><strong>Peak concurrent</strong> is the exact maximum number of sessions
  open simultaneously, with the moment it occurred.</p>
  <p><strong>95th percentile concurrent</strong> is the level exceeded in only
  5% of the 15-minute samples taken across the whole window &mdash; nights,
  weekends and holidays included, not 5% of working time. A raw peak is
  frequently a single anomalous spike, so this percentile is a steadier
  measure of typical concurrent load. Because the quiet hours are counted, it
  sits below the business-hours 95th percentile shown alongside it; if what
  matters is the working day rather than the average hour, read the
  business-hours column.</p>
  <p><strong>Anonymous sessions</strong> are reported separately because they do
  not consume a named user licence.</p>
  <p><strong>Unattributed sessions</strong> are sessions with no user identity
  recorded that are not flagged anonymous &mdash; a de-provisioned account or a
  broker glitch, in practice. They likewise do not consume a named user
  licence, and are reported separately from anonymous sessions so that every
  session in the window is accounted for: sessions belonging to an identified
  user, plus anonymous sessions, plus unattributed sessions, add up to the
  total sessions figure. <em>Unique users</em> is not a term in that sum
  &mdash; it counts people, not sessions, and one person routinely has many
  sessions in a window.</p>
  <p><strong>Business hours</strong> figures restrict the same measure to the
  15-minute samples that fall inside the working day shown beside them
  &mdash; by default 08:00&ndash;18:00, Monday to Friday, in the site's own
  time zone ($(ConvertTo-HtmlEncoded -Text $tz)), with daylight saving
  applied, rather than the whole week. Use it when the contract must be sized
  to cover the working day itself: the all-hours figure blends in nights,
  weekends and any activity outside that window, and whether that makes it
  higher, lower, or about the same as the business-hours figure depends on
  when this particular site is actually used &mdash; it is not always one
  direction, so read whichever column matches what the contract needs to
  cover rather than assuming one is simply the other plus a margin.</p>
  <p class="footnote">Generated by the Citrix Usage Report script. All times are
  shown in $(ConvertTo-HtmlEncoded -Text $tz) unless stated otherwise;
  concurrency is computed in UTC and business hours are evaluated in that same
  displayed time zone.</p>
</div>

</div>

<script>
  // Progressive enhancement only: the button is always visible, but does
  // nothing until this runs. The "or press Ctrl+P" hint next to it (see
  // .print-hint) covers the case where script does not run at all, so the
  // button is never a dead control with no way forward.
  (function () {
    var btn = document.getElementById('print-btn');
    if (btn) {
      btn.addEventListener('click', function () { window.print(); });
    }
  })();
</script>
</body>
</html>
"@
}

function Save-HtmlReport {
    <#
    .SYNOPSIS
        Renders and writes the report to disk.
    .DESCRIPTION
        Creates its output directory and writes the file with -ErrorAction
        Stop, then confirms the file actually exists before logging success.
        This does not rely solely on a caller having set
        $ErrorActionPreference = 'Stop' (as Invoke-CitrixUsageAudit does):
        a function that reports "Report written" must have checked, so that
        an unwritable path -- a nonexistent drive, a read-only or
        access-denied location -- can never produce a confident success
        message with nothing on disk, regardless of what error preference
        happens to be in effect wherever this is called from.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)][pscustomobject] $Analysis,
        [Parameter(Mandatory)][string] $Path
    )

    $dir = Split-Path -Parent $Path
    if ($dir -and -not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force -ErrorAction Stop | Out-Null
    }

    $html = New-HtmlReport -Analysis $Analysis
    Set-Content -Path $Path -Value $html -Encoding UTF8 -ErrorAction Stop

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Save-HtmlReport wrote no error but '$Path' does not exist afterwards; the report was not actually saved."
    }

    Write-AuditLog -Level Success -Message "Report written to $Path"
    return $Path
}

# endregion 75-HtmlReport.ps1

# ----------------------------------------------------------------------------
# region 80-Export.ps1
# ----------------------------------------------------------------------------
# ============================================================================
#  Export and anonymization
#
#  Anonymisation replaces identities with stable pseudonyms and never changes
#  a count. The mapping back to real identities is written to a separate file
#  that stays with the customer, so a privacy-sensitive site can still take
#  part in an audit without sending usernames anywhere.
# ============================================================================

function New-AnonymizationMap {
    <#
    .SYNOPSIS
        Builds a stable user id to pseudonym map for one dataset.
    .DESCRIPTION
        Built from the UNION of Users.Id and Sessions.UserId, not from Users
        alone. A session can reference a user id that the Users lookup does
        not carry (a de-provisioned account dropped from the lookup between
        fetches, for example) -- if the map were built from Users only, that
        session's UserId would fall through anonymisation untouched and a
        raw internal identifier would leak into an otherwise-anonymised
        export. Blank/whitespace UserIds (unattributed sessions) are not
        assigned a pseudonym; there is no identity there to protect.

        Users (and session-referenced ids) are sorted before numbering so the
        same dataset always yields the same map. An unstable map would make
        two runs of identical data disagree, and would stop an earlier
        identity map from decoding an earlier report.
    #>
    [CmdletBinding()]
    [OutputType([hashtable])]
    param([Parameter(Mandatory)][pscustomobject] $Dataset)

    $ids = @{}
    foreach ($u in $Dataset.Users) {
        if (-not [string]::IsNullOrWhiteSpace([string]$u.Id)) { $ids[[string]$u.Id] = $true }
    }
    foreach ($s in $Dataset.Sessions) {
        if (-not [string]::IsNullOrWhiteSpace([string]$s.UserId)) { $ids[[string]$s.UserId] = $true }
    }

    $map = @{}
    $index = 1
    foreach ($id in @($ids.Keys | Sort-Object)) {
        $map[$id] = 'User-{0:D4}' -f $index
        $index++
    }
    return $map
}

function New-IdentityMapRows {
    <#
    .SYNOPSIS
        The rows written to identity-map.csv: the customer's decode key.
    .DESCRIPTION
        The anonymisation map is pseudonym -> Citrix's internal user id, and
        that id is an opaque surrogate key (a GUID-like integer on a real
        site). Writing only those two columns produced a "decode key" that
        decodes nothing a person can read, while the runbook promised it
        mapped each pseudonym back to the real username and full name.

        So the real identity fields are joined in here, from the Users lookup.
        This MUST be called while $Dataset still holds real identities -- that
        is, after New-AnonymizationMap and BEFORE ConvertTo-AnonymizedDataset,
        which is exactly the window the orchestrator uses. Called after
        anonymisation it would faithfully write pseudonym -> pseudonym.

        Ids present in the map but absent from Users (a de-provisioned account
        referenced only by a session -- see New-AnonymizationMap) still get a
        row, with the name columns blank. Dropping them would leave a
        pseudonym in the report that the map cannot explain at all.
    #>
    [CmdletBinding()]
    [OutputType([object[]])]
    param(
        [Parameter(Mandatory)][hashtable] $Map,
        [AllowNull()][AllowEmptyCollection()] $Users
    )

    $byId = @{}
    foreach ($u in $Users) {
        $uid = [string] $u.Id
        if (-not [string]::IsNullOrWhiteSpace($uid)) { $byId[$uid] = $u }
    }

    $rows = foreach ($id in @($Map.Keys | Sort-Object)) {
        $u = $byId[[string] $id]
        [pscustomobject]@{
            Pseudonym  = $Map[$id]
            RealUserId = $id
            UserName   = $(if ($u) { [string] $u.UserName } else { '' })
            FullName   = $(if ($u) { [string] $u.FullName } else { '' })
            Upn        = $(if ($u) { [string] $u.Upn } else { '' })
        }
    }

    # Plain array return; callers wrap in @(). See the array-convention note
    # in src/60-Analytics.ps1.
    return @($rows)
}

function ConvertTo-AnonymizedDataset {
    <#
    .SYNOPSIS
        Returns a copy of the dataset with identities replaced.
    .DESCRIPTION
        Usernames, full names, UPNs, SIDs, domains, client device names,
        client addresses and session-level UserId are all replaced with
        stable pseudonyms. Session keys, timings, session type and the
        anonymous flag are left untouched.

        Mapping Sessions.UserId (not just Users) is required, not optional:
        Get-UniqueUserStats and every other session-derived breakdown reads
        UserId straight off Sessions, so leaving Sessions alone would let a
        raw user id resurface anywhere built from session data, defeating
        the anonymisation this function exists to provide.

        This does not change any count. The map built by
        New-AnonymizationMap is a bijection -- no two users share a
        pseudonym (see the dedicated test in tests/Export.Tests.ps1) -- so
        substituting UserId 1:1 through it preserves every unique-user and
        session count exactly, the same way substituting Id on Users does.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][pscustomobject] $Dataset,
        [Parameter(Mandatory)][hashtable] $Map
    )

    $users = foreach ($u in $Dataset.Users) {
        $alias = $Map[[string]$u.Id]
        if (-not $alias) { $alias = 'User-UNKNOWN' }
        [pscustomobject]@{
            Id = $u.Id; UserName = $alias; FullName = $alias
            Upn = "$alias@anonymized.local"; Sid = $null; Domain = 'ANONYMIZED'
        }
    }

    $sessions = foreach ($s in $Dataset.Sessions) {
        $rawUserId = [string]$s.UserId

        # Assigned to a plain local variable, not through a $(if...)
        # subexpression -- see the note below on ClientName/ClientAddress for
        # why that shape is avoided throughout this file. This value is
        # always a scalar (a string), never an array, so it is not exposed
        # to the collapse that pattern risks elsewhere, but the plain-
        # variable shape is used anyway for consistency.
        $mappedUserId = $rawUserId
        if (-not [string]::IsNullOrWhiteSpace($rawUserId)) {
            $alias = $Map[$rawUserId]
            if (-not $alias) { $alias = 'User-UNKNOWN' }
            $mappedUserId = $alias
        }

        [pscustomobject]@{
            SessionKey  = $s.SessionKey
            UserId      = $mappedUserId
            MachineId   = $s.MachineId
            StartUtc    = $s.StartUtc
            EndUtc      = $s.EndUtc
            SessionType = $s.SessionType
            IsAnonymous = $s.IsAnonymous
        }
    }

    # Device names and client addresses are pseudonymised with a per-value
    # counter rather than a hash, so the output is readable and the device
    # (or address) count is still verifiable by eye. Each gets its own
    # counter and map -- collapsing every address onto a single constant
    # value (as an earlier version of this function did) would silently
    # merge every distinct address into one bucket and corrupt
    # Get-ClientDeviceBreakdown's UniqueAddresses count, which is exactly
    # the figure that supports device-based licensing.
    $deviceMap = @{}
    $deviceIndex = 1
    $addressMap = @{}
    $addressIndex = 1
    $connections = foreach ($c in $Dataset.Connections) {
        $name = [string]$c.ClientName

        # Assigned to plain local variables, not through a $(if...)
        # subexpression: that shape re-enumerates its output and unwraps a
        # single-element array to a scalar even with @() already inside the
        # if -- see the array-convention note in src/60-Analytics.ps1's
        # Invoke-AuditAnalysis. These values are always scalars (a string
        # or $null), not arrays, so the collapse this file's callers must
        # guard against elsewhere cannot bite here -- but the plain-variable
        # shape is used anyway, for consistency with the one true pattern.
        $clientName = $null
        if ($name) {
            if (-not $deviceMap.ContainsKey($name)) {
                $deviceMap[$name] = 'Device-{0}' -f $deviceIndex
                $deviceIndex++
            }
            $clientName = $deviceMap[$name]
        }

        $clientAddress = $null
        $address = [string]$c.ClientAddress
        if ($address) {
            if (-not $addressMap.ContainsKey($address)) {
                $addressMap[$address] = 'Address-{0}' -f $addressIndex
                $addressIndex++
            }
            $clientAddress = $addressMap[$address]
        }

        [pscustomobject]@{
            Id = $c.Id; SessionKey = $c.SessionKey
            ClientName = $clientName
            ClientAddress = $clientAddress
            ClientVersion = $c.ClientVersion
            IsReconnect = $c.IsReconnect
        }
    }

    $copy = $Dataset.PSObject.Copy()
    $copy.Users = @($users)
    $copy.Sessions = @($sessions)
    $copy.Connections = @($connections)
    return $copy
}

function ConvertTo-AuditRoundTripUtcText {
    <#
    .SYNOPSIS
        Formats a UTC datetime unambiguously for a customer-facing CSV.
    .DESCRIPTION
        Formatted with an explicit invariant culture, not the machine's
        locale: a session-level export is exactly the file a customer in a
        different locale ends up opening on their own machine, and a
        locale-formatted timestamp (day/month order, non-ASCII digits, a
        different separator) would leave them with a file nobody -- not
        even Excel on a differently-configured machine -- can parse
        unambiguously. $null (a session with no end date, still running)
        maps to an empty string rather than a placeholder value, so the
        column stays genuinely blank instead of encoding "still running" as
        a magic string a reader has to know to interpret.

        The parameter is deliberately untyped rather than
        [Nullable[datetime]]: PowerShell 5.1's parameter binder does not
        reliably wrap a bound [datetime] value in a real Nullable<DateTime>
        (the bound value keeps its plain [datetime] runtime type), so
        `$Value.Value` silently returns $null -- accessing a nonexistent
        property does not error -- and `.ToString()` on that $null then
        throws "cannot call a method on a null-valued expression". Comparing
        against $null directly and calling .ToString() on $Value itself
        works for both a real [datetime] and an actual $null.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([AllowNull()] $Value)

    if ($null -eq $Value) { return '' }
    return $Value.ToString('yyyy-MM-ddTHH:mm:ssZ', [System.Globalization.CultureInfo]::InvariantCulture)
}

function Export-AuditData {
    <#
    .SYNOPSIS
        Writes the analysis and raw session data to JSON and CSV for
        independent re-analysis.
    .DESCRIPTION
        The exported JSON is a deliberately reduced view of the analysis
        object -- Analysis.Config carries the credential objects the run was
        authenticated with (ClientSecret, Credential), and those must never
        persist to disk. $exportable is built field by field rather than
        serialising $Analysis wholesale so a future field added to Config
        cannot silently leak into an exported file.

        sessions.csv is the raw, per-session data the dataset carries --
        distinct from every other file this function writes, which is
        derived/aggregated analysis. Its purpose is letting the customer
        re-analyse on their own side or spot-check the aggregate numbers.
        When -Dataset was produced by ConvertTo-AnonymizedDataset, its
        UserId column is already pseudonymous, because anonymisation is
        fixed at the dataset level (Sessions.UserId), not patched on at
        export time.
    .OUTPUTS
        The paths written, as a plain array. Callers must wrap the call in
        @() -- see the array-convention note in src/60-Analytics.ps1. This
        function deliberately does NOT `return ,$array`: a unary-comma-wrapped
        return only ever "looks" correct for a caller that captures a bare
        (non-@()-wrapped) call with exactly one file written, and is silently
        wrong -- Count comes back 1 regardless of the real number of files --
        for every other case once the caller applies the codebase's own @()
        convention on top of it. Verified empirically before writing this
        comment (see the "writes exactly one file" test in
        tests/Export.Tests.ps1).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][pscustomobject] $Analysis,
        [Parameter(Mandatory)][pscustomobject] $Dataset,
        [Parameter(Mandatory)][string] $OutputPath,
        [hashtable] $Map,

        # Built by New-IdentityMapRows BEFORE the dataset was anonymised, so
        # it still carries real usernames. $Dataset by this point holds
        # pseudonyms, so the rows cannot be derived here.
        [AllowNull()][AllowEmptyCollection()] $MapRows
    )

    if (-not (Test-Path $OutputPath)) {
        New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
    }

    $written = New-Object System.Collections.Generic.List[string]
    $windows = @($Analysis.Windows)
    $sessions = @($Dataset.Sessions)

    # Plain local with @() applied directly, never a `$(if ...)` inside the
    # hashtable literal below -- see the array-convention note in
    # src/60-Analytics.ps1. A single fetch warning is the likeliest case, and
    # a collapsed one would serialise to a bare object instead of a
    # one-element array in data.json.
    $fetchWarnings = @()
    if ($Analysis.PSObject.Properties['FetchWarnings']) { $fetchWarnings = @($Analysis.FetchWarnings) }

    # The full analysis, minus the config (which holds credential objects).
    $exportable = [pscustomobject]@{
        GeneratedUtc = $Analysis.GeneratedUtc
        Environment  = $Analysis.Config.EnvironmentLabel
        IsDemo       = $Analysis.IsDemo
        Preflight    = $Analysis.Preflight
        # Incomplete/unbounded fetches, so a re-analysis in the customer's own
        # tooling can see that a figure is a lower bound rather than a count.
        FetchWarnings = $fetchWarnings
        SiteTotals   = $Analysis.SiteTotals
        Windows      = @($windows | ForEach-Object {
            [pscustomobject]@{
                Days = $_.Days; WindowStartUtc = $_.WindowStartUtc; WindowEndUtc = $_.WindowEndUtc
                IsTruncated = $_.IsTruncated; UniqueUsers = $_.UniqueUsers
                AnonymousSessions = $_.AnonymousSessions
                # UnattributedSessions (sessions with a blank UserId that are
                # not flagged IsAnonymous) is carried through here so the
                # figures reconcile: attributed sessions + AnonymousSessions +
                # UnattributedSessions equals TotalSessions. (UniqueUsers is
                # not a term in that sum -- it counts people, not sessions.)
                # Omitting it would leave a silent gap in a customer-facing
                # report -- exactly the kind of unexplained number
                # Get-UniqueUserStats's own doc comment warns against.
                UnattributedSessions = $_.UnattributedSessions
                TotalSessions = $_.TotalSessions
                Concurrency = [pscustomobject]@{
                    Peak = $_.Concurrency.Peak; PeakAtUtc = $_.Concurrency.PeakAtUtc
                    P50 = $_.Concurrency.P50; P90 = $_.Concurrency.P90
                    P95 = $_.Concurrency.P95; P99 = $_.Concurrency.P99
                    BusinessP95 = $_.Concurrency.BusinessP95; BusinessPeak = $_.Concurrency.BusinessPeak
                    # What "business hours" meant for this run. Without it a
                    # reader re-analysing data.json cannot tell whether a
                    # BusinessP95 was measured over a London working day or a
                    # Sydney one, and the two are not comparable.
                    BusinessHourStart = $_.Concurrency.BusinessHourStart
                    BusinessHourEnd = $_.Concurrency.BusinessHourEnd
                    BusinessTimeZoneId = $_.Concurrency.BusinessTimeZoneId
                }
                DeliveryGroups = $_.DeliveryGroups
                SessionTypes   = $_.SessionTypes
                Applications   = $_.Applications
                ClientDevices  = $_.ClientDevices
                DailyTrend     = $_.DailyTrend
            }
        })
    }

    # Resolved to an absolute path first: WriteAllText below does not honour
    # PowerShell's current location the way Set-Content does, and the
    # directory has already been created above if it did not exist.
    $jsonDir = (Resolve-Path -LiteralPath $OutputPath).Path
    $jsonPath = Join-Path $jsonDir 'data.json'
    $json = $exportable | ConvertTo-Json -Depth 8

    # RFC 8259 Section 8.1: a JSON text must not begin with a byte order
    # mark. Set-Content -Encoding UTF8 always emits one on PowerShell 5.1,
    # which standard JSON tooling outside PowerShell (Python's json module,
    # most JavaScript parsers) rejects outright with an error that gives no
    # hint about the cause -- and data.json exists specifically so an
    # operator can re-analyse a customer's numbers in their own tooling.
    # Deliberately NOT applied to the CSVs below: Excel uses the BOM to
    # detect UTF-8, and those files are meant to be opened by a person, not
    # parsed by a standards-conforming machine reader. Same encoding
    # question, different correct answer, because the consumers differ.
    [System.IO.File]::WriteAllText($jsonPath, $json, (New-Object System.Text.UTF8Encoding($false)))
    [void] $written.Add($jsonPath)

    # Skipped when there are no windows to summarise, rather than writing an
    # empty, header-less CSV: Export-Csv still creates a zero-byte file for
    # zero pipeline input, and a file with nothing usable in it is worse than
    # no file at all.
    if ($windows.Count -gt 0) {
        $summaryPath = Join-Path $OutputPath 'summary.csv'
        @($windows | ForEach-Object {
            [pscustomobject]@{
                WindowDays              = $_.Days
                Truncated               = $_.IsTruncated
                UniqueUsers             = $_.UniqueUsers
                AnonymousSessions       = $_.AnonymousSessions
                UnattributedSessions    = $_.UnattributedSessions
                TotalSessions           = $_.TotalSessions
                PeakConcurrent          = $_.Concurrency.Peak
                P95Concurrent           = [math]::Round($_.Concurrency.P95, 1)
                BusinessHoursP95        = [math]::Round($_.Concurrency.BusinessP95, 1)
                BusinessHoursDefinition  = '{0:00}:00-{1:00}:00 Mon-Fri {2}' -f `
                    [int] $_.Concurrency.BusinessHourStart, [int] $_.Concurrency.BusinessHourEnd,
                    $_.Concurrency.BusinessTimeZoneId
            }
        }) | Export-Csv -Path $summaryPath -NoTypeInformation -Encoding UTF8
        [void] $written.Add($summaryPath)
    }

    $widest = $windows | Sort-Object Days -Descending | Select-Object -First 1
    if ($widest -and $widest.DailyTrend) {
        $trendPath = Join-Path $OutputPath 'daily-trend.csv'
        $widest.DailyTrend | Export-Csv -Path $trendPath -NoTypeInformation -Encoding UTF8
        [void] $written.Add($trendPath)
    }

    # Raw, per-session data -- distinct from every file above, which is
    # derived/aggregated analysis. Skipped (like summary.csv above) when
    # there is nothing to write, rather than emitting a useless empty file.
    # Timestamps are written in an explicit invariant round-trip UTC form
    # (see ConvertTo-AuditRoundTripUtcText), not the machine's locale
    # format, so a customer in another locale gets back a file they can
    # actually parse.
    if ($sessions.Count -gt 0) {
        $sessionsPath = Join-Path $OutputPath 'sessions.csv'
        $sessionRows = foreach ($s in $sessions) {
            [pscustomobject]@{
                SessionKey  = $s.SessionKey
                UserId      = $s.UserId
                MachineId   = $s.MachineId
                StartUtc    = ConvertTo-AuditRoundTripUtcText -Value $s.StartUtc
                EndUtc      = ConvertTo-AuditRoundTripUtcText -Value $s.EndUtc
                SessionType = $s.SessionType
                IsAnonymous = $s.IsAnonymous
            }
        }
        @($sessionRows) | Export-Csv -Path $sessionsPath -NoTypeInformation -Encoding UTF8
        [void] $written.Add($sessionsPath)
    }

    if ($Map) {
        # Written separately and named plainly, so the runbook can tell the
        # customer to send every file except this one.
        $mapPath = Join-Path $OutputPath 'identity-map.csv'
        $rows = @($MapRows)
        if ($rows.Count -eq 0) {
            # No pre-anonymisation snapshot was supplied, so only the internal
            # ids can be written. The name columns are still emitted, blank,
            # so the file's shape does not silently change between callers.
            $rows = @(New-IdentityMapRows -Map $Map -Users @())
        }
        $rows | Export-Csv -Path $mapPath -NoTypeInformation -Encoding UTF8
        [void] $written.Add($mapPath)
        Write-AuditLog -Level Warn -Message "Identity map written to $mapPath. Keep this file. Do not send it with the report."
    }

    Write-AuditLog -Level Success -Message "Exported $($written.Count) data file(s) to $OutputPath"

    # Plain array return, NOT `return ,$written.ToArray()`. This codebase's
    # convention (see src/60-Analytics.ps1) is: collection-returning functions
    # return a plain array and callers wrap the call in @(). A unary-comma
    # return instead makes the single object crossing the pipeline an array
    # THAT WRAPS the real array as its own lone element; a caller's @() then
    # wraps that a second time, and .Count silently reports 1 no matter how
    # many files were actually written (verified empirically: with two real
    # elements, @(fn-that-does-`return ,$arr`).Count comes back 1, not 2).
    return $written.ToArray()
}

# endregion 80-Export.ps1

# ----------------------------------------------------------------------------
# region 90-Gui.ps1
# ----------------------------------------------------------------------------
# ============================================================================
#  Interactive UI
#
#  A WinForms dialog where WinForms is available, guided console prompts
#  everywhere else. The dialog gathers configuration and returns it; it holds
#  no business logic, so the script behaves identically however it was
#  configured.
# ============================================================================

function Test-GuiAvailable {
    <#
    .SYNOPSIS
        Whether a WinForms dialog can be shown in this host.
    .DESCRIPTION
        Server Core, PowerShell 7 without the Windows compatibility layer, and
        non-interactive hosts all lack WinForms. Detection must never itself be
        the failure, so any error means "no GUI" and the console path is used.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param()

    if ($env:CITRIXAUDIT_FORCE_CONSOLE) { return $false }

    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
        Add-Type -AssemblyName System.Drawing -ErrorAction Stop
        return [System.Environment]::UserInteractive
    } catch {
        return $false
    }
}

function ConvertTo-DaysArray {
    <#
    .SYNOPSIS
        Parses a comma-separated window list into sorted unique integers.
    .DESCRIPTION
        Non-numeric entries, and zero or negative day counts, are ignored
        rather than rejected: a typo should not make the operator start over.

        Values above 1000 days are also dropped. 1000 is comfortably above
        the maximum Citrix Monitor retention on a Premium licence (90 days),
        so nothing legitimate is lost by the ceiling -- but with no ceiling
        at all, a single stray extra digit (-Days 99999) used to be accepted
        silently: the log would report "Windows: 99999 day(s)", demo
        generation would hang indefinitely trying to build a dataset that
        large, and against a real Citrix site it would issue a punishing
        OData query instead, with no validation error anywhere to explain
        why. Unlike the non-numeric/non-positive cases above, which are
        obvious noise a typo produces, an out-of-range value is real input
        the operator actually typed, so it is not dropped silently: a
        warning names the value and the bound, so a mistaken 30,60,99999
        visibly becomes 30,60 instead of quietly and misleadingly becoming
        30,60,90.
    #>
    [CmdletBinding()]
    [OutputType([int[]])]
    param([AllowEmptyString()][string] $Text)

    if ([string]::IsNullOrWhiteSpace($Text)) { return @(30, 60, 90) }

    $maxDays = 1000
    $droppedForRange = New-Object System.Collections.Generic.List[int]

    $values = foreach ($part in ($Text -split ',')) {
        $trimmed = $part.Trim()
        $n = 0
        if ([int]::TryParse($trimmed, [ref] $n) -and $n -gt 0) {
            if ($n -le $maxDays) {
                $n
            } else {
                [void] $droppedForRange.Add($n)
            }
        }
    }

    if ($droppedForRange.Count -gt 0) {
        $droppedList = ($droppedForRange | Sort-Object -Unique) -join ', '
        Write-AuditLog -Level Warn -Message "Ignored reporting window(s) over the $maxDays-day maximum: $droppedList."
    }

    $result = @($values | Sort-Object -Unique)
    if ($result.Count -eq 0) { return @(30, 60, 90) }
    return $result
}

function Get-GuiEnvironmentName {
    <#
    .SYNOPSIS
        Maps an environment dropdown/menu index to a Citrix environment name.
    .DESCRIPTION
        The environment picker (both the WinForms dropdown and the console
        menu) presents four choices in a fixed order: on-premises, Citrix
        Cloud Commercial, Citrix Cloud Japan, and Citrix Cloud Government.
        Government is a separate sovereign endpoint, so an off-by-one here
        would silently misroute credentials to the wrong cloud. That risk is
        why this mapping lives in its own small, testable function rather
        than as an inline array literal inside the dialog's click handler,
        where it would only ever be exercised by launching the GUI by hand.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [ValidateRange(0, 3)]
        [int] $Index
    )

    $map = @('OnPremises', 'CloudCommercial', 'CloudJapan', 'CloudGovernment')
    return $map[$Index]
}

function Get-CappedFormClientHeight {
    <#
    .SYNOPSIS
        The form's target ClientSize height, capped to fit the screen.
    .DESCRIPTION
        Pulled out of Show-AuditGui's dialog construction so the fit-to-
        screen arithmetic is testable without driving a modal dialog -- the
        same reason Get-GuiEnvironmentName and New-InteractiveCredential
        exist in this file. This is also the exact arithmetic that, worked
        out by hand while building the fix, first landed a form 29px taller
        than the working area (see the caller): a pure function that a test
        can pin down precisely is worth more here than a comment.

        $DesiredClientHeight and $Chrome must both be measured off a form
        whose native window handle already exists (Form.Height and
        Form.ClientSize.Height are unreliable estimates before that, and
        ClientSize can itself be silently reconciled to a slightly smaller
        value once the handle is created) -- see the caller for where that
        happens. This function only does the arithmetic once those two
        numbers are trustworthy.
    #>
    [CmdletBinding()]
    [OutputType([int])]
    param(
        [Parameter(Mandatory)][int] $DesiredClientHeight,
        [Parameter(Mandatory)][int] $Chrome,
        [Parameter(Mandatory)][int] $WorkingAreaHeight,
        [Parameter(Mandatory)][int] $MinimumClientHeight,
        [int] $ScreenMargin = 60
    )

    $maxFormHeight = $WorkingAreaHeight - $ScreenMargin
    $formHeight = $DesiredClientHeight + $Chrome
    if ($formHeight -le $maxFormHeight) { return $DesiredClientHeight }

    $overage = $formHeight - $maxFormHeight
    return [Math]::Max($MinimumClientHeight, $DesiredClientHeight - $overage)
}

function New-InteractiveCredential {
    <#
    .SYNOPSIS
        Builds a PSCredential from the interactive on-premises auth choice,
        or returns $null to signal "authenticate as the signed-in user".
    .DESCRIPTION
        Both the WinForms dialog and the console fallback offer the same
        choice: run as the signed-in user (the default), or supply a
        different Windows account. This is the decision behind that choice,
        pulled out of both click handlers/prompts so it is testable without
        driving a modal dialog -- the same reason Get-GuiEnvironmentName
        exists in this file. $null flows straight to New-AuditConfig
        -Credential, which is what makes Get-AuthRequestParameters
        (src/30-Auth.ps1) fall back to -UseDefaultCredentials.
    #>
    [CmdletBinding()]
    [OutputType([System.Management.Automation.PSCredential])]
    param(
        [bool] $UseDifferentAccount,
        [AllowEmptyString()][AllowNull()][string] $UserName,
        [System.Security.SecureString] $Password
    )

    if (-not $UseDifferentAccount) { return $null }

    if ([string]::IsNullOrWhiteSpace($UserName)) {
        throw 'A username is required when "Use a different account" is selected.'
    }

    # A masked password field left empty is a legitimate, if unusual, choice
    # -- not a reason to crash. PSCredential's constructor requires a
    # non-null SecureString, so an empty one stands in for "no password".
    $pw = $Password
    if (-not $pw) { $pw = New-Object System.Security.SecureString }

    return [System.Management.Automation.PSCredential]::new($UserName, $pw)
}

function Show-AuditGui {
    <#
    .SYNOPSIS
        Shows the configuration dialog and returns a config object.
    .OUTPUTS
        The configuration, or $null if the operator cancelled.
    #>
    [CmdletBinding()]
    param([string] $DefaultOutputPath)

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $p = Get-BrandPalette
    function HexColor([string] $Hex) {
        $h = $Hex.TrimStart('#')
        [System.Drawing.Color]::FromArgb(
            [Convert]::ToInt32($h.Substring(0,2),16),
            [Convert]::ToInt32($h.Substring(2,2),16),
            [Convert]::ToInt32($h.Substring(4,2),16))
    }

    $jasper = HexColor $p['Jasper80']
    $cobalt = HexColor $p['Cobalt50']
    $canvas = HexColor $p['Canvas20']
    $mint   = HexColor $p['Mint40']

    # Funnel Display and Public Sans are usually absent on a customer machine,
    # so the dialog asks for them and lets GDI fall back to a system font.
    $headFont = New-Object System.Drawing.Font('Funnel Display', 15, [System.Drawing.FontStyle]::Bold)
    $bodyFont = New-Object System.Drawing.Font('Public Sans', 9)

    # -----------------------------------------------------------------
    # Layout: three regions stacked in the form, only the middle one scrolls.
    #
    #   +----------------------------------------------------+
    #   | header (Dock=Top)     -- branding, never scrolls   |
    #   +----------------------------------------------------+
    #   | content (Dock=Fill,   -- inputs; AutoScroll=$true  |
    #   |          AutoScroll)     grows/shrinks with the     |
    #   |                           window, scrolls when the  |
    #   |                           content is taller than    |
    #   |                           the visible area          |
    #   +----------------------------------------------------+
    #   | action bar (Dock=Bottom) -- Run audit / Close,      |
    #   |                              always visible          |
    #   +----------------------------------------------------+
    #
    # Docking order matters and is counter to the usual mental model here:
    # verified empirically (see docs referenced in the task-24 report) that
    # in this WinForms runtime the Fill-docked control must be the FIRST one
    # added to the form's Controls collection, with the Top- and
    # Bottom-docked controls added after it. Adding Fill last (or between
    # Top and Bottom) reliably produces overlap -- the fill panel claims the
    # full client area at the moment it's added and does not yield space
    # back to a Top/Bottom control added afterward. This was confirmed with
    # a standalone repro before being relied on here, not assumed from the
    # framework's documentation, which is commonly misremembered the other
    # way round.
    $headerHeight = 78
    $actionBarHeight = 64
    $formWidth = 640

    $content = New-Object System.Windows.Forms.Panel
    $content.Dock = 'Fill'
    $content.AutoScroll = $true
    $content.BackColor = $canvas

    $header = New-Object System.Windows.Forms.Panel
    $header.Dock = 'Top'
    $header.Height = $headerHeight
    $header.BackColor = $jasper

    $actionBar = New-Object System.Windows.Forms.Panel
    $actionBar.Dock = 'Bottom'
    $actionBar.Height = $actionBarHeight
    $actionBar.BackColor = $canvas

    # Fill first, then Top, then Bottom -- see the note above.
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Citrix Usage Report'
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'Sizable'
    $form.MaximizeBox = $true
    $form.BackColor = $canvas
    $form.Font = $bodyFont
    $form.Controls.Add($content)
    $form.Controls.Add($header)
    $form.Controls.Add($actionBar)

    # Give the form its real target width right away, before any anchored
    # control below is created. A control's Anchor distance to its parent's
    # right edge is captured against the parent's size at the moment the
    # control is added/anchored -- and a freshly-created Form starts at
    # WinForms' own default size (around 300px wide), not $formWidth. Every
    # Left+Right (or Top,Right-only) anchored control below would otherwise
    # bake in its distance against that narrow default, then balloon far
    # past the form's actual right edge the first time the height is set
    # below. Confirmed by reproducing exactly that ballooning before adding
    # this line. Height is a placeholder here; only the width needs to be
    # final this early, and the height set later (which does not change the
    # width) leaves these horizontal anchors alone.
    $form.ClientSize = New-Object System.Drawing.Size($formWidth, 700)

    $title = New-Object System.Windows.Forms.Label
    $title.Text = 'Citrix Usage Report'
    $title.ForeColor = [System.Drawing.Color]::White
    $title.Font = $headFont
    $title.Location = New-Object System.Drawing.Point(24, 14)
    $title.AutoSize = $true
    $header.Controls.Add($title)

    $subtitle = New-Object System.Windows.Forms.Label
    $subtitle.Text = 'Unique users and concurrent sessions over the requested periods'
    $subtitle.ForeColor = $mint
    $subtitle.Location = New-Object System.Drawing.Point(26, 46)
    $subtitle.AutoSize = $true
    $header.Controls.Add($subtitle)

    # Everything below is added to the scrolling content panel, not the
    # form, so $y is relative to the top of that panel rather than the top
    # of the form (the header used to occupy the first 96px of this
    # coordinate space; now it's a separate docked region).
    $y = 20
    function New-Label([string] $Text, [int] $Top) {
        $l = New-Object System.Windows.Forms.Label
        $l.Text = $Text
        $l.Location = New-Object System.Drawing.Point(24, $Top)
        $l.AutoSize = $true
        $l
    }

    $content.Controls.Add((New-Label 'Environment' $y))
    $y += 22
    $envBox = New-Object System.Windows.Forms.ComboBox
    $envBox.DropDownStyle = 'DropDownList'
    $envBox.Location = New-Object System.Drawing.Point(24, $y)
    $envBox.Size = New-Object System.Drawing.Size(570, 24)
    $envBox.Anchor = 'Top, Left, Right'
    # Order matters: it must line up exactly with Get-GuiEnvironmentName.
    [void] $envBox.Items.AddRange(@(
        'On-premises CVAD (Delivery Controller)',
        'Citrix Cloud - Commercial (US / EU / Asia Pacific South)',
        'Citrix Cloud - Japan',
        'Citrix Cloud Government (US Gov)'
    ))
    $envBox.SelectedIndex = 0
    $content.Controls.Add($envBox)
    $y += 36

    $content.Controls.Add((New-Label 'Delivery Controller hostname' $y))
    $y += 22
    $serverBox = New-Object System.Windows.Forms.TextBox
    $serverBox.Location = New-Object System.Drawing.Point(24, $y)
    $serverBox.Size = New-Object System.Drawing.Size(420, 24)
    # Grows with the form; protocolBox (below) tracks the right edge instead
    # so the two stay a fixed gap apart rather than the hostname box either
    # clipping or leaving a growing gap before the dropdown as the form is
    # resized.
    $serverBox.Anchor = 'Top, Left, Right'
    # The Monitor Service can be published over either scheme -- that is an
    # IIS setting on the Delivery Controller, not something this tool can
    # infer -- so the protocol is an explicit choice sitting right next to
    # the hostname it applies to, defaulting to the recommended Https.
    $protocolBox = New-Object System.Windows.Forms.ComboBox
    $protocolBox.DropDownStyle = 'DropDownList'
    $protocolBox.Location = New-Object System.Drawing.Point(452, $y)
    $protocolBox.Size = New-Object System.Drawing.Size(142, 24)
    $protocolBox.Anchor = 'Top, Right'
    [void] $protocolBox.Items.AddRange(@('Https (recommended)', 'Http (unencrypted)'))
    $protocolBox.SelectedIndex = 0
    # An explicit scheme typed into the hostname (a pasted browser URL, or a
    # customer who already worked out their Monitor Service is HTTP-only)
    # wins over whatever the dropdown was last set to, and the dropdown
    # updates to say so rather than silently disagreeing with what was typed.
    $serverBox.Add_TextChanged({
        if ($serverBox.Text -match '^\s*(https?)://') {
            $protocolBox.SelectedIndex = if ($matches[1] -eq 'http') { 1 } else { 0 }
        }
    })
    $content.Controls.AddRange(@($serverBox, $protocolBox))
    $y += 36

    # Previously one shared caption ("Customer ID / Service Principal ID /
    # Service Principal Secret") over three boxes side by side with an 8px
    # gap between them -- a user had to map a position in the caption to a
    # position in the row, and with the gap that small the three boxes read
    # as one control. Stacked here instead, each field with its own label
    # directly above it (the same label-above-box pattern every other field
    # in this dialog uses, e.g. Environment/Delivery Controller hostname
    # above), so identifying a field never depends on counting columns. The
    # scrolling content panel (see the layout note at the top of this
    # function) is what makes the extra vertical room affordable.
    $customerLabel = New-Label 'Citrix Cloud Customer ID' $y
    $content.Controls.Add($customerLabel)
    $y += 22
    $customerBox = New-Object System.Windows.Forms.TextBox
    $customerBox.Location = New-Object System.Drawing.Point(24, $y)
    $customerBox.Size = New-Object System.Drawing.Size(300, 24)
    $content.Controls.Add($customerBox)
    $y += 36

    $clientIdLabel = New-Label 'Service Principal ID' $y
    $content.Controls.Add($clientIdLabel)
    $y += 22
    $clientIdBox = New-Object System.Windows.Forms.TextBox
    $clientIdBox.Location = New-Object System.Drawing.Point(24, $y)
    $clientIdBox.Size = New-Object System.Drawing.Size(300, 24)
    $content.Controls.Add($clientIdBox)
    $y += 36

    $secretLabel = New-Label 'Service Principal Secret' $y
    $content.Controls.Add($secretLabel)
    $y += 22
    $secretBox = New-Object System.Windows.Forms.TextBox
    $secretBox.Location = New-Object System.Drawing.Point(24, $y)
    $secretBox.Size = New-Object System.Drawing.Size(300, 24)
    $secretBox.UseSystemPasswordChar = $true
    $content.Controls.Add($secretBox)
    $y += 36

    # Citrix renamed "Secure Clients" to "Service principals" in the console;
    # the field labels above follow that, and this link takes the operator
    # straight to Citrix's guidance for creating one, since "go find it
    # yourself" is a poor answer to hand someone mid-dialog. LinkLabel opens
    # the OS default browser on click; Start-Process can fail if a machine
    # has no default browser association, and a dead link is a nuisance but
    # an unhandled exception out of a dialog's click handler is worse, so
    # that failure is caught and reported in the status label instead.
    $cloudDocsUrl = 'https://developer-docs.citrix.com/en-us/citrix-cloud/citrix-cloud-api-overview/get-started-with-citrix-cloud-apis.html'
    $cloudDocsLink = New-Object System.Windows.Forms.LinkLabel
    $cloudDocsLink.Text = 'Citrix docs: configuring a Citrix Cloud service principal'
    $cloudDocsLink.Location = New-Object System.Drawing.Point(24, $y)
    $cloudDocsLink.Size = New-Object System.Drawing.Size(420, 20)
    $cloudDocsLink.AutoSize = $true
    $cloudDocsLink.LinkColor = $cobalt
    $cloudDocsLink.ActiveLinkColor = $cobalt
    $cloudDocsLink.VisitedLinkColor = $cobalt
    $cloudDocsLink.LinkBehavior = 'HoverUnderline'
    $cloudDocsLink.Add_LinkClicked({
        try {
            Start-Process $cloudDocsUrl | Out-Null
        } catch {
            $status.ForeColor = [System.Drawing.Color]::Firebrick
            $status.Text = "Couldn't open a browser. Documentation: $cloudDocsUrl"
        }
    })
    $content.Controls.Add($cloudDocsLink)
    $y += 30

    # On-premises authentication choice. Defaults to the signed-in user (the
    # long-standing behaviour); "Use a different account" reveals a username
    # and masked password field, mirrored by Read-AuditConfigFromConsole's
    # console prompt below so the two entry points offer the same choice.
    # Single column, stacked, not a side-by-side pair: see the note above
    # the "Include in the report" GroupBox about why this file avoids a
    # column layout for check-style controls (a RadioButton renders its
    # selection glyph the same way a CheckBox does).
    $authLabel = New-Label 'Run as (on-premises only)' $y
    $content.Controls.Add($authLabel)
    $y += 22

    $rbCurrentUser = New-Object System.Windows.Forms.RadioButton
    $rbCurrentUser.Text = 'Run as the signed-in user'
    $rbCurrentUser.Location = New-Object System.Drawing.Point(24, $y)
    $rbCurrentUser.Size = New-Object System.Drawing.Size(400, 22)
    $rbCurrentUser.Checked = $true
    $content.Controls.Add($rbCurrentUser)
    $y += 24

    $rbDifferentAccount = New-Object System.Windows.Forms.RadioButton
    $rbDifferentAccount.Text = 'Use a different account'
    $rbDifferentAccount.Location = New-Object System.Drawing.Point(24, $y)
    $rbDifferentAccount.Size = New-Object System.Drawing.Size(400, 22)
    $content.Controls.Add($rbDifferentAccount)
    $y += 28

    $acctUserLabel = New-Label 'Username (DOMAIN\user or user@domain)' $y
    $content.Controls.Add($acctUserLabel)
    $y += 22
    $acctUserBox = New-Object System.Windows.Forms.TextBox
    $acctUserBox.Location = New-Object System.Drawing.Point(24, $y)
    $acctUserBox.Size = New-Object System.Drawing.Size(300, 24)
    $content.Controls.Add($acctUserBox)
    $y += 30

    $acctPassLabel = New-Label 'Password' $y
    $content.Controls.Add($acctPassLabel)
    $y += 22
    $acctPassBox = New-Object System.Windows.Forms.TextBox
    $acctPassBox.Location = New-Object System.Drawing.Point(24, $y)
    $acctPassBox.Size = New-Object System.Drawing.Size(300, 24)
    $acctPassBox.UseSystemPasswordChar = $true
    $content.Controls.Add($acctPassBox)
    $y += 36

    # Cloud fields are meaningless on-premises and vice versa, so only the
    # relevant set is enabled; the account fields are further gated on
    # "Use a different account" actually being selected.
    $updateFields = {
        $isCloud = $envBox.SelectedIndex -gt 0
        $serverBox.Enabled = -not $isCloud
        $protocolBox.Enabled = -not $isCloud
        $customerLabel.Enabled = $isCloud
        $customerBox.Enabled = $isCloud
        $clientIdLabel.Enabled = $isCloud
        $clientIdBox.Enabled = $isCloud
        $secretLabel.Enabled = $isCloud
        $secretBox.Enabled = $isCloud
        $cloudDocsLink.Enabled = $isCloud

        $authLabel.Enabled = -not $isCloud
        $rbCurrentUser.Enabled = -not $isCloud
        $rbDifferentAccount.Enabled = -not $isCloud

        $useDifferentAccount = (-not $isCloud) -and $rbDifferentAccount.Checked
        $acctUserLabel.Enabled = $useDifferentAccount
        $acctUserBox.Enabled = $useDifferentAccount
        $acctPassLabel.Enabled = $useDifferentAccount
        $acctPassBox.Enabled = $useDifferentAccount
    }
    $envBox.Add_SelectedIndexChanged($updateFields)
    $rbCurrentUser.Add_CheckedChanged($updateFields)
    $rbDifferentAccount.Add_CheckedChanged($updateFields)
    & $updateFields

    $content.Controls.Add((New-Label 'Reporting windows, in days' $y))
    $y += 22
    $daysBox = New-Object System.Windows.Forms.TextBox
    $daysBox.Text = '30,60,90'
    $daysBox.Location = New-Object System.Drawing.Point(24, $y)
    $daysBox.Size = New-Object System.Drawing.Size(180, 24)
    $content.Controls.Add($daysBox)
    $y += 36

    $content.Controls.Add((New-Label 'Output folder' $y))
    $y += 22
    $outBox = New-Object System.Windows.Forms.TextBox
    $outBox.Text = $DefaultOutputPath
    $outBox.Location = New-Object System.Drawing.Point(24, $y)
    $outBox.Size = New-Object System.Drawing.Size(490, 24)
    $outBox.Anchor = 'Top, Left, Right'
    $browse = New-Object System.Windows.Forms.Button
    $browse.Text = 'Browse'
    $browse.Location = New-Object System.Drawing.Point(520, ($y - 1))
    $browse.Size = New-Object System.Drawing.Size(74, 26)
    $browse.Anchor = 'Top, Right'
    $browse.Add_Click({
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        if ($dlg.ShowDialog() -eq 'OK') { $outBox.Text = $dlg.SelectedPath }
    })
    $content.Controls.AddRange(@($outBox, $browse))
    $y += 40

    $group = New-Object System.Windows.Forms.GroupBox
    $group.Text = 'Include in the report'
    $group.Location = New-Object System.Drawing.Point(24, $y)
    $group.Size = New-Object System.Drawing.Size(570, 206)
    $group.Anchor = 'Top, Left, Right'

    # Single column, not the two-column grid this used to be: at x=340 (the
    # right-hand column), the second and third rows consistently rendered
    # with no check glyph at all - reproduced in isolation with a bare
    # GroupBox and confirmed to depend only on that (x, row) slot, not on
    # which control, text, or Checked value occupied it. A single column
    # never landed a control in that slot in any repro, so it is the
    # reliable layout rather than a narrower one-off pixel tweak.
    $cbGroups = New-Object System.Windows.Forms.CheckBox
    $cbGroups.Text = 'Delivery group breakdown'
    $cbGroups.Location = New-Object System.Drawing.Point(16, 26)
    $cbGroups.Size = New-Object System.Drawing.Size(500, 22)
    $cbGroups.Checked = $true

    $cbApps = New-Object System.Windows.Forms.CheckBox
    $cbApps.Text = 'Session types and published applications (slower)'
    $cbApps.Location = New-Object System.Drawing.Point(16, 50)
    $cbApps.Size = New-Object System.Drawing.Size(500, 22)

    $cbDevices = New-Object System.Windows.Forms.CheckBox
    $cbDevices.Text = 'Client devices and Workspace app versions (slower)'
    $cbDevices.Location = New-Object System.Drawing.Point(16, 74)
    $cbDevices.Size = New-Object System.Drawing.Size(500, 22)

    $cbTrend = New-Object System.Windows.Forms.CheckBox
    $cbTrend.Text = 'Daily activity trend'
    $cbTrend.Location = New-Object System.Drawing.Point(16, 98)
    $cbTrend.Size = New-Object System.Drawing.Size(500, 22)
    $cbTrend.Checked = $true

    $cbAnon = New-Object System.Windows.Forms.CheckBox
    $cbAnon.Text = 'Anonymize usernames'
    $cbAnon.Location = New-Object System.Drawing.Point(16, 122)
    $cbAnon.Size = New-Object System.Drawing.Size(500, 22)

    $cbExport = New-Object System.Windows.Forms.CheckBox
    $cbExport.Text = 'Export raw data (JSON + CSV)'
    $cbExport.Location = New-Object System.Drawing.Point(16, 146)
    $cbExport.Size = New-Object System.Drawing.Size(500, 22)
    $cbExport.Checked = $true

    $cbDemo = New-Object System.Windows.Forms.CheckBox
    $cbDemo.Text = 'Demo mode (synthetic data)'
    $cbDemo.Location = New-Object System.Drawing.Point(16, 170)
    $cbDemo.Size = New-Object System.Drawing.Size(500, 22)

    $group.Controls.AddRange(@($cbGroups, $cbApps, $cbDevices, $cbTrend, $cbAnon, $cbExport, $cbDemo))
    $content.Controls.Add($group)
    $y += 222

    $status = New-Object System.Windows.Forms.Label
    $status.Location = New-Object System.Drawing.Point(24, $y)
    $status.Size = New-Object System.Drawing.Size(570, 20)
    $status.Anchor = 'Top, Left, Right'
    $status.ForeColor = HexColor $p['Jasper80']
    $content.Controls.Add($status)
    $y += 26

    # $y now measures the full height the content panel needs to lay out
    # every input control without clipping or cropping anything -- this is
    # the "content height" the caller reasons about when sizing the form.
    # AutoScrollMinSize makes that explicit for the panel: whatever doesn't
    # fit in the panel's visible height becomes scrollable rather than cut
    # off, no matter how the form itself ends up sized below.
    $contentPadBottom = 16
    $contentHeight = $y + $contentPadBottom
    $content.AutoScrollMinSize = New-Object System.Drawing.Size(0, $contentHeight)

    # Cap the form's height to the working area so it always fits the
    # display -- a tall fixed form on a laptop or a modest-resolution RDP
    # session is exactly how the buttons ended up off-screen in the first
    # place. ClientSize is set once to the natural (uncapped) size so the
    # actual non-client chrome (title bar, borders -- both DPI-dependent)
    # can be measured directly off the real Form.Height rather than
    # estimated, then reduced if that total would exceed the working area.
    #
    # This happens BEFORE the action bar's buttons are created and anchored
    # below, deliberately: a control's Anchor distance is fixed relative to
    # its parent's size at the moment the anchor takes effect, and actionBar
    # is docked Bottom, so its width tracks the form's ClientSize.Width.
    # Positioning/anchoring the buttons against the form's *eventual* width
    # -- after this resize, not before -- is what keeps them inside the
    # action bar instead of drifting off its right edge the first time the
    # form is resized (verified: doing this in the other order was tried
    # and visibly broke, with the buttons landing well past the bar's right
    # edge).
    $desiredClientHeight = $headerHeight + $contentHeight + $actionBarHeight
    $form.ClientSize = New-Object System.Drawing.Size($formWidth, $desiredClientHeight)

    # Force the form's native window handle into existence before measuring
    # it (this does not show the window). Two things are only accurate once
    # the handle exists:
    #  - Form.Height is otherwise an underestimate, short by roughly the
    #    title bar's height, because the real non-client metrics aren't
    #    resolved until there's a real HWND (observed: 10px "chrome" before
    #    handle creation vs. the real ~39px at 100% DPI).
    #  - ClientSize itself can be silently reconciled downward slightly once
    #    the handle exists (observed: a ~30px shrink from the value just
    #    assigned above). The height-cap arithmetic below must measure
    #    against this real, reconciled ClientSize -- not the value that was
    #    requested before the handle existed -- or the cap undershoots and
    #    the final form still exceeds the working area.
    [void] $form.Handle
    $actualClientHeight = $form.ClientSize.Height
    $chrome = $form.Height - $actualClientHeight

    # Resizable rather than a fixed dialog -- see the header comment on
    # $form.FormBorderStyle above. The minimum keeps every field's fixed
    # x-position from overlapping (the widest rows -- the environment
    # dropdown and the "Include in the report" group box -- run out to
    # x=594) and keeps the action bar plus a usable slice of content on
    # screen even when the user shrinks the window; it also doubles as the
    # floor Get-CappedFormClientHeight won't shrink below.
    $minUsableContent = 220
    $minimumClientHeight = $headerHeight + $minUsableContent + $actionBarHeight

    $cappedClientHeight = Get-CappedFormClientHeight -DesiredClientHeight $actualClientHeight `
        -Chrome $chrome `
        -WorkingAreaHeight ([System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea.Height) `
        -MinimumClientHeight $minimumClientHeight
    if ($cappedClientHeight -ne $actualClientHeight) {
        $form.ClientSize = New-Object System.Drawing.Size($formWidth, $cappedClientHeight)
    }

    $form.MinimumSize = New-Object System.Drawing.Size(
        ($formWidth + 20),
        ($minimumClientHeight + $chrome))

    # Action bar: fixed, docked to the bottom of the form, outside the
    # scrolling content panel, so Run audit / Close are visible regardless
    # of scroll position or window size. Positioned against actionBar's
    # actual (now-final) width rather than the $formWidth constant, so this
    # still lines up correctly even if the height cap above also changed
    # the form's width story in some future edit.
    $barWidth = $actionBar.ClientSize.Width
    $run = New-Object System.Windows.Forms.Button
    $run.Text = 'Run Report'
    $run.Location = New-Object System.Drawing.Point(($barWidth - 24 - 110 - 8 - 76), 15)
    $run.Size = New-Object System.Drawing.Size(110, 34)
    $run.Anchor = 'Top, Right'
    $run.BackColor = $cobalt
    $run.ForeColor = [System.Drawing.Color]::White
    $run.FlatStyle = 'Flat'
    $run.FlatAppearance.BorderSize = 0

    $cancel = New-Object System.Windows.Forms.Button
    $cancel.Text = 'Close'
    $cancel.Location = New-Object System.Drawing.Point(($barWidth - 24 - 76), 15)
    $cancel.Size = New-Object System.Drawing.Size(76, 34)
    $cancel.Anchor = 'Top, Right'
    $cancel.FlatStyle = 'Flat'

    $actionBar.Controls.AddRange(@($run, $cancel))

    $script:GuiResult = $null

    $run.Add_Click({
        try {
            $secure = $null
            if ($secretBox.Text) {
                $secure = ConvertTo-SecureString -String $secretBox.Text -AsPlainText -Force
            }

            # Becomes a SecureString immediately, same as the cloud client
            # secret above, and is never written to disk or logged. $null
            # here (the "signed-in user" radio) is what makes New-AuditConfig
            # leave -Credential unset, so Get-AuthRequestParameters falls
            # back to -UseDefaultCredentials.
            $acctSecure = $null
            if ($acctPassBox.Text) {
                $acctSecure = ConvertTo-SecureString -String $acctPassBox.Text -AsPlainText -Force
            }
            $credential = New-InteractiveCredential -UseDifferentAccount $rbDifferentAccount.Checked `
                -UserName $acctUserBox.Text -Password $acctSecure

            $selectedProtocol = if ($protocolBox.SelectedIndex -eq 1) { 'Http' } else { 'Https' }

            $cfg = New-AuditConfig `
                -Environment (Get-GuiEnvironmentName -Index $envBox.SelectedIndex) `
                -DeliveryController $serverBox.Text `
                -Protocol $selectedProtocol `
                -CustomerId $customerBox.Text `
                -ClientId $clientIdBox.Text `
                -ClientSecret $secure `
                -Credential $credential `
                -Days (ConvertTo-DaysArray -Text $daysBox.Text) `
                -OutputPath $outBox.Text `
                -IncludeDeliveryGroups:$cbGroups.Checked `
                -IncludeApplications:$cbApps.Checked `
                -IncludeClientDevices:$cbDevices.Checked `
                -IncludeTrend:$cbTrend.Checked `
                -Anonymize:$cbAnon.Checked `
                -ExportRawData:$cbExport.Checked `
                -DemoData:$cbDemo.Checked

            # Validate before closing, so a mistake is corrected here rather
            # than surfacing as a stack trace after the window has gone.
            Test-AuditConfig -Config $cfg

            $script:GuiResult = $cfg
            $form.Close()
        } catch {
            $status.ForeColor = [System.Drawing.Color]::Firebrick
            $status.Text = $_.Exception.Message
        }
    })

    $cancel.Add_Click({ $script:GuiResult = $null; $form.Close() })

    [void] $form.ShowDialog()
    return $script:GuiResult
}

function Read-AuditConfigFromConsole {
    <#
    .SYNOPSIS
        Guided console prompts, for hosts without WinForms.
    #>
    [CmdletBinding()]
    param([string] $DefaultOutputPath)

    Write-Host ''
    Write-Host '  Citrix Usage Report' -ForegroundColor Cyan
    Write-Host '  Unique users and concurrent sessions over the requested periods'
    Write-Host ''

    # Defined up front, ahead of the on-premises auth choice below, which is
    # the earliest of its several uses in this function.
    function Confirm-Toggle([string] $Prompt, [bool] $Default) {
        $suffix = if ($Default) { '[Y/n]' } else { '[y/N]' }
        $answer = Read-Host "  $Prompt $suffix"
        if (-not $answer) { return $Default }
        return $answer -match '^(y|yes)$'
    }

    Write-Host '  1. On-premises CVAD (Delivery Controller)'
    Write-Host '  2. Citrix Cloud - Commercial (US / EU / Asia Pacific South)'
    Write-Host '  3. Citrix Cloud - Japan'
    Write-Host '  4. Citrix Cloud Government (US Gov)'
    Write-Host ''

    $choice = Read-Host '  Environment [1]'
    if (-not $choice) { $choice = '1' }
    $choiceIndex = 0
    if (-not [int]::TryParse($choice, [ref] $choiceIndex) -or $choiceIndex -lt 1 -or $choiceIndex -gt 4) {
        $choiceIndex = 1
    }
    # Menu options are numbered 1-4; Get-GuiEnvironmentName is 0-based, and is
    # the single source of truth shared with the dialog's dropdown so the two
    # entry points can never disagree about which choice means Government.
    $environment = Get-GuiEnvironmentName -Index ($choiceIndex - 1)

    $server = ''; $protocol = 'Https'; $customerId = ''; $clientId = ''; $secret = $null; $credential = $null

    if ($environment -eq 'OnPremises') {
        $server = Read-Host '  Delivery Controller hostname'

        # An explicit scheme typed into the hostname wins over the prompt
        # below and sets its default, the same way the dialog's dropdown
        # follows what was typed into the hostname box.
        $detectedScheme = $null
        if ($server -match '^\s*(https?)://') { $detectedScheme = $matches[1].ToLowerInvariant() }
        $defaultProtocol = if ($detectedScheme -eq 'http') { 'Http' } else { 'Https' }

        $protocolAnswer = Read-Host "  Monitor Service protocol, the Delivery Controller's IIS setting -- Https or Http [$defaultProtocol]"
        $protocol = if (-not $protocolAnswer) { $defaultProtocol }
            elseif ($protocolAnswer -match '^(?i)http$') { 'Http' }
            else { 'Https' }

        # Mirrors the dialog's "Run as" radio buttons: default to the
        # signed-in user, offer a different account as the alternative. The
        # password is read straight into a SecureString via -AsSecureString
        # -- it is never captured as plain text.
        if (Confirm-Toggle 'Use a different account instead of the signed-in user?' $false) {
            $acctUserName = Read-Host '  Username (DOMAIN\user or user@domain)'
            $acctPassword = Read-Host '  Password' -AsSecureString
            $credential = New-InteractiveCredential -UseDifferentAccount $true `
                -UserName $acctUserName -Password $acctPassword
        }
    } else {
        # Citrix renamed "Secure Clients" to "Service principals" in the
        # console; the prompts below follow that, and the URL is printed as
        # plain text because a console can't render a link the way the
        # dialog's LinkLabel does -- a customer on Server Core only ever
        # sees this path.
        Write-Host '  Create a service principal under Identity and Access Management > API Access > Service principals in the Citrix Cloud console.'
        Write-Host '  Guide: https://developer-docs.citrix.com/en-us/citrix-cloud/citrix-cloud-api-overview/get-started-with-citrix-cloud-apis.html'
        Write-Host ''
        $customerId = Read-Host '  Citrix Cloud Customer ID'
        $clientId   = Read-Host '  Service Principal ID'
        $secret     = Read-Host '  Service Principal Secret' -AsSecureString
    }

    $daysText = Read-Host '  Reporting windows in days [30,60,90]'
    $outPath  = Read-Host "  Output folder [$DefaultOutputPath]"
    if (-not $outPath) { $outPath = $DefaultOutputPath }

    $cfg = New-AuditConfig -Environment $environment -DeliveryController $server -Protocol $protocol `
        -CustomerId $customerId -ClientId $clientId -ClientSecret $secret -Credential $credential `
        -Days (ConvertTo-DaysArray -Text $daysText) -OutputPath $outPath `
        -IncludeDeliveryGroups:(Confirm-Toggle 'Delivery group breakdown?' $true) `
        -IncludeApplications:(Confirm-Toggle 'Published application breakdown (slower)?' $false) `
        -IncludeClientDevices:(Confirm-Toggle 'Client device breakdown (slower)?' $false) `
        -IncludeTrend:(Confirm-Toggle 'Daily activity trend?' $true) `
        -Anonymize:(Confirm-Toggle 'Anonymize usernames?' $false) `
        -ExportRawData:(Confirm-Toggle 'Export raw data?' $true) `
        -DemoData:(Confirm-Toggle 'Demo mode (synthetic data, no Citrix environment contacted)?' $false)

    Test-AuditConfig -Config $cfg
    return $cfg
}

# endregion 90-Gui.ps1

# ----------------------------------------------------------------------------
# region 99-Main.ps1
# ----------------------------------------------------------------------------
# ============================================================================
#  Orchestration
#
#  config -> preflight -> fetch -> analyse -> render -> export
# ============================================================================

function Invoke-CitrixUsageAudit {
    <#
    .SYNOPSIS
        Runs a complete audit.
    .OUTPUTS
        An exit code: 0 success, 1 failure, 2 cancelled.
    #>
    [CmdletBinding()]
    [OutputType([int])]
    param([hashtable] $BoundParameters = @{})

    # Every cmdlet failure anywhere in this call chain must reach the catch
    # block below, which is the only place a customer-facing message and a
    # non-zero exit code get produced. Left at the default, most cmdlet
    # failures (New-Item, Set-Content, Join-Path against a missing drive,
    # ...) are non-terminating: they print a raw error to the console and
    # execution continues past them as if nothing had happened -- which is
    # how an unwritable -OutputPath used to produce a confident "Audit
    # complete", real-looking numbers, and exit 0, with nothing written to
    # disk. Assigned to a local (function-scoped) variable, not $global:, so
    # it cannot leak into a caller's session -- e.g. the test suite, which
    # dot-sources this file and must not have its own preference silently
    # overwritten. Write-AuditLog's own internal try/catch around its log
    # write already sets -ErrorAction Stop explicitly and catches it itself,
    # so logging still degrades to console-only on a write failure rather
    # than becoming the reason a run fails, regardless of this setting.
    $ErrorActionPreference = 'Stop'

    $stamp = [DateTime]::Now.ToString('yyyyMMdd-HHmmss')

    try {
        # ---- Configuration -------------------------------------------------
        $defaultOut = if ($BoundParameters.ContainsKey('OutputPath')) {
            $BoundParameters['OutputPath']
        } else {
            Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'CitrixUsageReport'
        }

        $interactive = -not ($BoundParameters.ContainsKey('DeliveryController') -or
                             $BoundParameters.ContainsKey('CustomerId') -or
                             $BoundParameters.ContainsKey('DemoData'))

        if ($interactive) {
            # In the interactive path the configuration comes wholly from the
            # dialog (or the console prompts), so anything else the caller
            # bound on the command line is thrown away. Doing that silently
            # produced a report that looked fine and simply was not the run
            # the operator asked for -- a -Anonymize that never happened is
            # the worst case. OutputPath and NoGui are excluded because both
            # ARE honoured: OutputPath seeds the dialog's default, and NoGui
            # selects the console path.
            $ignorable = @('Environment','DeliveryController','Protocol','CustomerId','ClientId',
                           'ClientSecret','Credential','Days','IncludeDeliveryGroups',
                           'IncludeApplications','IncludeClientDevices','IncludeTrend',
                           'Anonymize','ExportRawData')
            $ignored = @($ignorable | Where-Object { $BoundParameters.ContainsKey($_) })
            if ($ignored.Count -gt 0) {
                Write-AuditLog -Level Warn -Message ("No connection details were supplied on the command line, so this run is being configured from the dialog and these parameter(s) are IGNORED: -{0}. Set them in the dialog instead, or add -DeliveryController / -CustomerId / -DemoData to run non-interactively." -f ($ignored -join ', -'))
            }

            # A WinForms dialog can only be shown from a single-threaded
            # apartment. Test-GuiAvailable (src/90-Gui.ps1) checks that
            # WinForms loads and that the host is interactive, but it does
            # not -- and must not, since it has no reason to know its
            # caller's threading model -- check apartment state. This is the
            # entry point that decides GUI vs. console, so the check belongs
            # here: calling Show-AuditGui from an MTA thread would pass
            # Test-GuiAvailable and then throw from ShowDialog() itself,
            # turning a should-be console fallback into a crash.
            $guiRequested = -not $BoundParameters['NoGui']
            $guiAvailable = $guiRequested -and (Test-GuiAvailable)

            if ($guiAvailable -and
                [System.Threading.Thread]::CurrentThread.GetApartmentState() -ne
                [System.Threading.ApartmentState]::STA) {
                Write-AuditLog -Level Warn -Message 'The graphical dialog requires a single-threaded apartment (STA); this session is not one. Falling back to console prompts.'
                $guiAvailable = $false
            }

            $config = if ($guiAvailable) {
                Show-AuditGui -DefaultOutputPath $defaultOut
            } else {
                Read-AuditConfigFromConsole -DefaultOutputPath $defaultOut
            }

            if (-not $config) {
                Write-Host 'Cancelled.' -ForegroundColor Yellow
                return 2
            }
        } else {
            # Pre-assigned to plain locals rather than threaded through a
            # $(if ...) subexpression: see the array-convention note in
            # src/60-Analytics.ps1's Invoke-AuditAnalysis. Days is the one
            # value here that is genuinely array-typed, so it is the one
            # genuinely at risk of the single-element collapse; the rest are
            # scalars and are pre-assigned only for consistency and
            # readability.
            $auditEnvironment = 'OnPremises'
            if ($BoundParameters['Environment']) { $auditEnvironment = $BoundParameters['Environment'] }

            # $Days is [string[]] on the command line (see src/00-Header.ps1
            # for why), not [int[]]: under `-File` -- the invocation the
            # runbook must recommend, since it is the one that works with
            # -ExecutionPolicy Bypass against a Restricted default policy --
            # PowerShell hands the script raw argv strings and parses none
            # of it as PowerShell syntax. `-Days 30,60,90` arrives as ONE
            # string "30,60,90" under -File, but as three separate strings
            # "30","60","90" when the same call is dot-sourced or invoked
            # via -Command (where a real array literal is parsed). Joining
            # whatever shape arrived back into one comma-separated string
            # before handing it to ConvertTo-DaysArray (src/90-Gui.ps1, the
            # same parser the GUI and console paths already use) normalises
            # both shapes -- and a single value -- to the same result:
            # @("30,60,90") -join ',' and @("30","60","90") -join ',' both
            # produce "30,60,90". ConvertTo-DaysArray itself is only defined
            # by the time this function runs, not while the param block
            # above is being bound, which is why the parsing happens here.
            $auditDays = @(30, 60, 90)
            if ($BoundParameters['Days']) {
                $auditDays = @(ConvertTo-DaysArray -Text (($BoundParameters['Days'] -join ',')))
            }

            $auditProtocol = 'Https'
            if ($BoundParameters['Protocol']) { $auditProtocol = [string] $BoundParameters['Protocol'] }

            $config = New-AuditConfig `
                -Environment $auditEnvironment `
                -DeliveryController ([string] $BoundParameters['DeliveryController']) `
                -Protocol $auditProtocol `
                -CustomerId ([string] $BoundParameters['CustomerId']) `
                -ClientId ([string] $BoundParameters['ClientId']) `
                -ClientSecret $BoundParameters['ClientSecret'] `
                -Credential $BoundParameters['Credential'] `
                -Days $auditDays `
                -OutputPath $defaultOut `
                -IncludeDeliveryGroups:([bool] $BoundParameters['IncludeDeliveryGroups']) `
                -IncludeApplications:([bool] $BoundParameters['IncludeApplications']) `
                -IncludeClientDevices:([bool] $BoundParameters['IncludeClientDevices']) `
                -IncludeTrend:([bool] $BoundParameters['IncludeTrend']) `
                -Anonymize:([bool] $BoundParameters['Anonymize']) `
                -ExportRawData:([bool] $BoundParameters['ExportRawData']) `
                -DemoData:([bool] $BoundParameters['DemoData'])

            Test-AuditConfig -Config $config
        }

        # ---- Output location and logging ------------------------------------
        # Validated and probed for real writability here, before any other
        # work (including Initialize-AuditLog and the fetch/generate step
        # below, which can run for minutes): a nonexistent drive or an
        # access-denied location must fail immediately with a message naming
        # the path, not silently continue into a run that ends in a
        # confident "Audit complete" with nothing on disk, and not surface
        # as some unrelated-looking error one statement further on (a
        # nonexistent drive letter makes Join-Path itself throw
        # "Cannot find drive", which -- before this path was validated up
        # front and before $ErrorActionPreference was set to 'Stop' above --
        # used to be swallowed as a non-terminating error, leaving $runDir
        # as an empty string that then made the *next* statement fail with
        # the unrelated-looking "Cannot bind argument to parameter 'Path'
        # because it is null"). Test-Path never throws even for an unmapped
        # drive letter (it returns $false), so it is safe to call before the
        # path is known to exist; the actual create-or-probe below is what
        # can fail, and is wrapped so its failure is rethrown with a message
        # a customer can act on.
        try {
            if (-not (Test-Path -LiteralPath $config.OutputPath)) {
                New-Item -ItemType Directory -Path $config.OutputPath -Force -ErrorAction Stop | Out-Null
            }
            # A directory that already exists is not necessarily writable
            # (a read-only location, a permissions issue) -- Test-Path only
            # confirms it is there. Probing with a real write-then-delete
            # catches that case too, not just "does not exist yet".
            $writeProbePath = Join-Path $config.OutputPath ".citrixaudit-write-test-$stamp"
            Set-Content -Path $writeProbePath -Value '' -Encoding UTF8 -ErrorAction Stop
            Remove-Item -LiteralPath $writeProbePath -Force -ErrorAction Stop
        } catch {
            throw "The output path '$($config.OutputPath)' could not be created or is not writable. Check that the drive exists and that this account can write to it. ($($_.Exception.Message))"
        }

        $runDir = Join-Path $config.OutputPath "CitrixUsageReport-$stamp"
        if (-not (Test-Path $runDir)) { New-Item -ItemType Directory -Path $runDir -Force -ErrorAction Stop | Out-Null }

        Initialize-AuditLog -Path (Join-Path $runDir 'usage-report.log')
        Write-AuditLog -Level Info -Message "Target: $($config.EnvironmentLabel)"
        Write-AuditLog -Level Info -Message "Windows: $($config.Days -join ', ') day(s)"

        # Which account authenticates is useful for diagnosing a 401 -- see
        # the account-aware message in Invoke-CitrixODataQuery (40-ODataClient.ps1).
        # The username is not sensitive and is safe to log; the password
        # never reaches this or any other log line.
        if (-not $config.IsCloud -and -not $config.DemoData) {
            if ($config.Credential) {
                Write-AuditLog -Level Info -Message "Authenticating on-premises as $($config.Credential.UserName)."
            } else {
                Write-AuditLog -Level Info -Message "Authenticating on-premises as the signed-in user ($env:USERDOMAIN\$env:USERNAME)."
            }
        }

        # HTTP was an explicit choice (-Protocol Http, or an http:// scheme
        # typed into -DeliveryController) -- there is no silent fallback to
        # warn about here, only the fact that this run is unencrypted. Kept
        # short; the HTML report carries the same notice with a clickable
        # link to Citrix's TLS guidance.
        if ($config.IsHttp) {
            Write-AuditLog -Level Warn -Message "This report was generated over an unencrypted HTTP connection to the Monitor Service. See Citrix's guidance on securing the Monitor Service with TLS: $script:MonitorTlsGuidanceUrl"
        }

        # ---- Collect ---------------------------------------------------------
        $progress = {
            param($entity, $count)
            Write-Progress -Activity 'Collecting Citrix usage data' `
                -Status "$entity : $count record(s)" -Id 1
        }

        $dataset = if ($config.DemoData) {
            Write-AuditLog -Level Warn -Message 'Demo mode: generating synthetic data. No Citrix environment will be contacted.'
            New-DemoDataset -Config $config
        } else {
            $auth = New-CitrixAuthContext -Config $config
            Get-CitrixDataset -AuthContext $auth -Config $config -ProgressAction $progress
        }

        Write-Progress -Activity 'Collecting Citrix usage data' -Id 1 -Completed

        if (@($dataset.Sessions).Count -eq 0) {
            Write-AuditLog -Level Warn -Message 'No sessions were returned for the requested period. The report will be empty. Check that the account can read Monitor data and that the site has had activity in this window.'
        }

        # ---- Anonymize -------------------------------------------------------
        $map = $null
        $mapRows = @()
        if ($config.Anonymize) {
            Write-AuditLog -Level Info -Message 'Anonymizing usernames...'
            $map = New-AnonymizationMap -Dataset $dataset
            # Snapshot the decode key BEFORE the dataset is anonymised: this
            # is the only moment $dataset.Users still holds real usernames,
            # and identity-map.csv is worthless to the customer without them.
            $mapRows = @(New-IdentityMapRows -Map $map -Users $dataset.Users)
            $dataset = ConvertTo-AnonymizedDataset -Dataset $dataset -Map $map
        }

        # ---- Analyze ---------------------------------------------------------
        Write-AuditLog -Level Info -Message 'Analyzing...'
        $analysis = Invoke-AuditAnalysis -Dataset $dataset -Config $config

        # ---- Render ----------------------------------------------------------
        $reportPath = Join-Path $runDir 'CitrixUsageReport.html'
        Save-HtmlReport -Analysis $analysis -Path $reportPath | Out-Null

        if ($config.ExportRawData) {
            Export-AuditData -Analysis $analysis -Dataset $dataset -OutputPath $runDir `
                -Map $map -MapRows $mapRows | Out-Null
        } elseif ($map) {
            # The identity map is still written even when raw data export is
            # off, so the operator can decode the (possibly anonymised)
            # report later. Export-AuditData is deliberately NOT called here:
            # it always writes the full raw export set (data.json,
            # summary.csv, daily-trend.csv, sessions.csv) alongside
            # identity-map.csv, with no way to ask for the map alone. Calling
            # it in this branch would hand the operator every raw file they
            # just declined by leaving -ExportRawData off, silently defeating
            # the toggle. Only identity-map.csv is written here, matching
            # Export-AuditData's own identity-map.csv logic exactly.
            $mapPath = Join-Path $runDir 'identity-map.csv'
            $mapRows | Export-Csv -Path $mapPath -NoTypeInformation -Encoding UTF8
            Write-AuditLog -Level Warn -Message "Identity map written to $mapPath. Keep this file. Do not send it with the report."
        }

        # ---- Summary ---------------------------------------------------------
        $widest = $analysis.Windows | Sort-Object Days -Descending | Select-Object -First 1

        Write-AuditLog -Level Success -Message ''
        Write-AuditLog -Level Success -Message 'Report complete.'
        Write-AuditLog -Level Info -Message "  Unique users ($($widest.Days) days): $($widest.UniqueUsers)"
        Write-AuditLog -Level Info -Message "  Peak concurrent          : $($widest.Concurrency.Peak)"
        Write-AuditLog -Level Info -Message "  p95 concurrent           : $([math]::Round($widest.Concurrency.P95, 0))"
        Write-AuditLog -Level Info -Message "  Report                   : $reportPath"

        if ($widest.IsTruncated) {
            Write-AuditLog -Level Warn -Message "  NOTE: the $($widest.Days)-day window is truncated to $($widest.AvailableDays) days of retained history. The figures are a lower bound."
        }

        if ($map) {
            Write-AuditLog -Level Warn -Message '  Keep identity-map.csv. Do not send it with the report.'
        }

        if ([Environment]::UserInteractive -and -not $BoundParameters['NoGui']) {
            try { Start-Process $reportPath } catch { }
        }

        return 0

    } catch {
        $message = $_.Exception.Message
        Write-AuditLog -Level Error -Message "Report failed: $message"
        Write-AuditLog -Level Debug -Message $_.ScriptStackTrace
        Write-Host ''
        Write-Host 'The report did not complete. See the message above and the run log for detail.' -ForegroundColor Red
        return 1
    }
}

# Entry point. Runs only when this file is executed as a script, never when it
# is dot-sourced by the test suite.
if ($MyInvocation.InvocationName -ne '.' -and -not $env:CITRIXAUDIT_SUPPRESS_MAIN) {
    exit (Invoke-CitrixUsageAudit -BoundParameters $PSBoundParameters)
}

# endregion 99-Main.ps1


