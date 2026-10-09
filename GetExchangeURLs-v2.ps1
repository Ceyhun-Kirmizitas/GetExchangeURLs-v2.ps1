<#
.SYNOPSIS
    Displays Exchange Client Access URLs and optional authentication settings.

.DESCRIPTION
    GetExchangeURLs-v2.ps1 is read-only. It reads Exchange Client Access
    configuration and does not change Exchange configuration.

    The script shows Client Access URLs and namespaces for Autodiscover, ECP,
    EWS, MAPI, ActiveSync, OAB, OWA, PowerShell, and Outlook Anywhere.

    Only Exchange servers whose role is Mailbox or ClientAccess are queried.
    Edge Transport servers are not queried.

    If -Server is omitted, all Exchange servers with the Mailbox or ClientAccess
    role are queried.
    If -Server is used, each name must resolve to exactly one Exchange server.
    Short names and FQDNs are accepted. A name that cannot be resolved is
    skipped with a warning and never expands the query to all servers.
    A selected server whose role is not Mailbox or ClientAccess, for example
    Edge Transport, is skipped with a warning.
    The same server is processed only once.

    If -Service is omitted, all supported Client Access services are displayed.

    Default virtual directory queries use -AdPropertiesOnly for faster URL
    collection. Use -IncludeAuthentication when authentication and IIS-backed
    settings are also required. Authentication properties that a component
    does not expose are not shown.

    Client Access Service output includes the Alternate Service Account (ASA)
    configuration to help identify Kerberos deployments. The script shows
    whether an ASA is configured and, when available, the account name.
    It does not test ASA credential health and never reads ASA passwords.
    If the ASA status cannot be read, the script reads the Client Access
    Service again without the ASA status and shows the ASA as Unknown.

    One failed query does not stop the other queries. Failed queries are shown
    as WARNING lines and are counted in the completion message.

.PARAMETER Server
    Optional. Exchange server(s) to query, by name or FQDN.
    Only servers with the Mailbox or ClientAccess role are queried.
    If omitted, all Exchange servers with the Mailbox or ClientAccess role
    are queried.

.PARAMETER Service
    Optional. Services to display: Autodiscover, ECP, EWS, MAPI, ActiveSync,
    OAB, OWA, PowerShell, OutlookAnywhere.

.PARAMETER GroupByService
    Optional. Groups selected results by service across Exchange servers.

.PARAMETER IncludeAuthentication
    Includes authentication and IIS-backed settings.

.PARAMETER OutputFile
    Saves the same formatted output to a text file.
    The output directory and write access are checked before Exchange is
    queried. A missing directory is created. An existing file is overwritten.

.EXAMPLE
    .\GetExchangeURLs-v2.ps1
    Shows all supported Client Access services on all Exchange servers with
    the Mailbox or ClientAccess role.

.EXAMPLE
    .\GetExchangeURLs-v2.ps1 -Server EX01
    Shows all supported Client Access services on EX01.

.EXAMPLE
    .\GetExchangeURLs-v2.ps1 -Service MAPI,EWS
    Shows only MAPI and EWS on all Exchange servers with the Mailbox or
    ClientAccess role.

.EXAMPLE
    .\GetExchangeURLs-v2.ps1 -Server EX01,EX02 -GroupByService
    Compares each service across EX01 and EX02.

.EXAMPLE
    .\GetExchangeURLs-v2.ps1 -Server EX01,EX02 -Service MAPI,EWS -GroupByService
    Compares MAPI and EWS across EX01 and EX02.

.EXAMPLE
    .\GetExchangeURLs-v2.ps1 -Server EX01 -IncludeAuthentication
    Shows URLs and authentication settings on EX01.

.EXAMPLE
    .\GetExchangeURLs-v2.ps1 -Server EX01 -OutputFile C:\Temp\ExchangeURLs.txt
    Shows the results on EX01 and saves the same output to a text file.

.NOTES
Author        : Ceyhun Kirmizitas
Version       : 2.2
Date          : 09/10/2026
Applies to    : Exchange servers with the Mailbox or ClientAccess role. Edge Transport is not supported.
Compatibility : Intended for Exchange Server 2016, Exchange Server 2019, and Exchange Server Subscription Edition.
Validated     : Live lab validation performed on Exchange build 15.2.1748.10 (Mailbox role) using Windows PowerShell 5.1.
                Configured ASA, ASA status failure, Edge Transport, and report save failure paths were validated offline only.
                Exchange Server 2016 and Exchange Server Subscription Edition were not tested.
Version gate  : None; no hard Exchange-version gate is enforced.
Shell         : Windows PowerShell 5.1
Mode          : Read-only. No Exchange configuration changes are made.
Website       : https://ceyhunkirmizitas.net
GitHub        : https://github.com/Ceyhun-Kirmizitas
LinkedIn      : https://www.linkedin.com/in/ceyhun-kirmizitas/

Credits
-------
Originally written by : Paul Cunningham
Edited by             : Ali Tajran
Further updated by    : Ceyhun Kirmizitas

License
-------
MIT License
Copyright (c) 2015 Paul Cunningham
Copyright (c) 2026 Ceyhun Kirmizitas

Caution
-------
Use this script at your own risk. Review and test it in your environment before production use.
The author is not responsible for any issues, outages, or data loss resulting from its use.

Change Log
----------
2.2 - 09/10/2026
      The Client Access Service and the Autodiscover virtual directory are
      queried separately. A failed ASA status query no longer hides the SCP
      URI or the Autodiscover virtual directory. If the ASA status cannot be
      read, the Client Access Service is read again without it and the ASA
      is shown as Unknown.
      -Server names must resolve to exactly one Exchange server by name or
      FQDN. A name that cannot be resolved is skipped with a warning and
      never expands the query to all servers. The same server is processed
      only once.
      A selected server whose role is not Mailbox or ClientAccess, for
      example Edge Transport, is skipped with a warning instead of being
      removed silently.
      Removed the ASA Credential Status line. It did not test credential
      health. The Alternate Service Account line shows the account name,
      Not Configured, or Unknown.
      -OutputFile: the output directory and write access are checked before
      Exchange is queried. A failed report write is shown as an ERROR, and
      "Output saved to" is shown only after the report was saved.
      New startup header, completion message, and Feedback / Bugs footer.
      The completion message shows when one or more queries failed.
      A "Reading current Exchange values" message is shown once before the
      Exchange queries start.

2.1 - 14/09/2026
      Added Alternate Service Account (ASA) visibility to the
      Client Access Service output.
      Replaced the previous GroupBy modes with the simpler
      -GroupByService switch for comparing the same service
      configuration across multiple Exchange servers.
      Simplified the script structure and removed unnecessary
      processing layers.
      Improved default query performance by using
      -AdPropertiesOnly where applicable.
      Improved text-file output handling.
      Updated by Ceyhun Kirmizitas.

2.0 - 25/08/2026
      Updated Get-ClientAccessServer to Get-ClientAccessService.
      Added Autodiscover virtual directory details.
      Added optional authentication output with -IncludeAuthentication.
      Added optional text-file output with -OutputFile.
      Added N/A and Not Configured output.
      Removed -ADPropertiesOnly so AD and IIS-backed properties
      can be returned where available.
      Made -Server optional; if omitted, all Exchange Mailbox
      servers in the organization are queried.
      Added -Service filtering (for example: -Service MAPI,EWS).
      Added -GroupBy Server and -GroupBy Service output modes.
      Added combined filtering/grouping support.
      Expanded built-in examples and parameter documentation.
      Updated by Ceyhun Kirmizitas.

1.10 - 18/04/2020
      Added PowerShell virtual directory and reordered output.
      Updated by Ali Tajran.

1.00 - 27/08/2015
      Initial version by Paul Cunningham.

.LINK
https://github.com/Ceyhun-Kirmizitas/GetExchangeURLs-v2.ps1

.LINK
https://ceyhunkirmizitas.net

.LINK
https://github.com/Ceyhun-Kirmizitas

.LINK
https://www.linkedin.com/in/ceyhun-kirmizitas/
#>

#requires -version 5.1

[CmdletBinding()]
param(
    [Parameter(Position=0)]
    [string[]]$Server,

    [ValidateSet("Autodiscover","ECP","EWS","MAPI","ActiveSync","OAB","OWA","PowerShell","OutlookAnywhere")]
    [string[]]$Service,

    [switch]$GroupByService,

    [switch]$IncludeAuthentication,

    [string]$OutputFile
)

$script:ScriptBaseName = "GetExchangeURLs-v2"
$script:ScriptVersion = "2.2"

# Output is buffered when -OutputFile is used, then written once at the end.
$script:OutputLines = New-Object 'System.Collections.Generic.List[string]'
$script:OutputTarget = $null
$script:WarningCount = 0

# ---------------------------------------------------------------------------
# Output helpers
# ---------------------------------------------------------------------------
function Write-Result
{
    param([string]$Text = "", [ConsoleColor]$Color)

    if ($PSBoundParameters.ContainsKey("Color"))
    {
        Write-Host $Text -ForegroundColor $Color
    }
    else
    {
        Write-Host $Text
    }

    if ($OutputFile)
    {
        [void]$script:OutputLines.Add($Text)
    }
}

# Every WARNING line is counted so the completion message can report
# partially failed queries.
function Write-ResultWarning
{
    param([string]$Text)

    $script:WarningCount++
    Write-Result "WARNING: $Text" Yellow
}

# Implementation reference: CK-GEURL-01
function Write-Field
{
    param([string]$Label, [object]$Value)

    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace("$Value"))
    {
        $Value = "Not Configured"
    }
    elseif ($Value -is [System.Collections.IEnumerable] -and -not ($Value -is [string]))
    {
        $Items = @($Value | ForEach-Object { "$_" })
        $Value = if ($Items.Count) { $Items -join ", " } else { "Not Configured" }
    }

    Write-Result (" - {0,-34}: {1}" -f $Label, $Value)
}

function Show-StartupBanner
{
    Write-Result ""
    Write-Result ("{0}.ps1" -f $script:ScriptBaseName)
    Write-Result "Exchange Server client access URL and authentication review" Cyan
    Write-Result ""
    Write-Result "Author  : Ceyhun Kirmizitas" Cyan
    Write-Result ("Version : {0}" -f $script:ScriptVersion) Cyan
    Write-Result "Mode    : Configuration Review" Cyan
    Write-Result "Changes : NONE" Green
    Write-Result ""
    Write-Result "The script reads Client Access URLs, ASA configuration, and optional authentication settings."
    Write-Result "No Exchange configuration changes will be made."
    Write-Result ""
    Write-Result "Compatibility note: Intended for Exchange Server 2016, Exchange Server 2019, and Exchange Server Subscription Edition." Yellow
    Write-Result "No hard Exchange-version gate is enforced." Yellow
    Write-Result ""
}

# Console only. The footer is never written to the TXT report.
function Show-CompletionFooter
{
    Write-Host ""
    Write-Host "############################################" -ForegroundColor Cyan
    Write-Host "# Feedback / Bugs" -ForegroundColor Cyan
    Write-Host "############################################" -ForegroundColor Cyan
    Write-Host "For updates, feedback, bugs, and feature requests:"
    Write-Host "https://github.com/Ceyhun-Kirmizitas" -ForegroundColor DarkGray
    Write-Host "https://ceyhunkirmizitas.net" -ForegroundColor DarkGray
    Write-Host "############################################" -ForegroundColor Cyan
}

# Controlled stop for conditions where the review cannot continue.
function Stop-Review
{
    param([string]$Message)

    Write-Host ""
    Write-Host "ERROR: $Message" -ForegroundColor Red
    Write-Host ""
    Write-Host "No Exchange configuration changes were made." -ForegroundColor DarkGray
    Show-CompletionFooter
    exit 1
}

# ---------------------------------------------------------------------------
# Output file
# ---------------------------------------------------------------------------
# Checked before Exchange is queried, so a long run does not fail only at the
# final write. The write test uses a separate temporary file and never touches
# the destination file content.
function Initialize-OutputFile
{
    param([string]$Path)

    $FullPath = $null
    try
    {
        $FullPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
    }
    catch
    {
        throw "-OutputFile '$Path' is not a valid file path. $($_.Exception.Message)"
    }

    if (Test-Path -LiteralPath $FullPath -PathType Container)
    {
        throw "-OutputFile '$FullPath' is an existing directory. Specify a file path."
    }

    $Parent = [System.IO.Path]::GetDirectoryName($FullPath)
    if ([string]::IsNullOrWhiteSpace($Parent))
    {
        throw "-OutputFile '$FullPath' does not resolve to a valid parent directory."
    }

    if (-not (Test-Path -LiteralPath $Parent -PathType Container))
    {
        try
        {
            New-Item -Path $Parent -ItemType Directory -Force -ErrorAction Stop | Out-Null
        }
        catch
        {
            throw "The output directory '$Parent' could not be created. $($_.Exception.Message)"
        }
    }

    $TestFile = Join-Path $Parent (".{0}-writetest-{1}.tmp" -f $script:ScriptBaseName, [guid]::NewGuid().ToString("N"))
    try
    {
        $Stream = [System.IO.File]::Open($TestFile, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
        $Stream.Dispose()
    }
    catch
    {
        throw "The output directory '$Parent' is not writable. $($_.Exception.Message)"
    }
    finally
    {
        if (Test-Path -LiteralPath $TestFile -PathType Leaf)
        {
            Remove-Item -LiteralPath $TestFile -Force -ErrorAction SilentlyContinue
        }
    }

    # An existing report is overwritten at the end. Opening it for write
    # without writing confirms this is possible and leaves the content as is.
    if (Test-Path -LiteralPath $FullPath -PathType Leaf)
    {
        try
        {
            $Stream = [System.IO.File]::Open($FullPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Write, [System.IO.FileShare]::ReadWrite)
            $Stream.Dispose()
        }
        catch
        {
            throw "The existing file '$FullPath' cannot be overwritten. $($_.Exception.Message)"
        }
    }

    return $FullPath
}

# UTF-8 with BOM and CRLF, the same format as Set-Content -Encoding UTF8 in
# Windows PowerShell 5.1.
function Save-OutputFile
{
    param([string]$Path)

    $Writer = New-Object System.IO.StreamWriter -ArgumentList @($Path, $false, (New-Object System.Text.UTF8Encoding -ArgumentList $true))
    try
    {
        foreach ($Line in $script:OutputLines)
        {
            $Writer.Write($Line)
            $Writer.Write("`r`n")
        }
    }
    finally
    {
        $Writer.Dispose()
    }
}

# ---------------------------------------------------------------------------
# Exchange server resolution
# ---------------------------------------------------------------------------
# Get-ExchangeServer -Identity is not trusted on its own. The returned object
# must match the requested short name or FQDN, so an unexpected result can
# never expand the query to other servers.
function Resolve-ExchangeServerName
{
    param([string]$Name)

    $Requested = $Name.Trim()
    $Found = $null

    try
    {
        $Found = @(Get-ExchangeServer -Identity $Requested -ErrorAction Stop)
    }
    catch
    {
        Write-ResultWarning "Exchange server '$Requested' was not found. $($_.Exception.Message)"
        return $null
    }

    $Match = @($Found | Where-Object { "$($_.Name)" -eq $Requested -or "$($_.Fqdn)" -eq $Requested })

    if ($Match.Count -ne 1)
    {
        Write-ResultWarning "'$Requested' did not resolve to exactly one Exchange server. Use the server name or FQDN; wildcards are not supported."
        return $null
    }

    return $Match[0]
}

function Get-ServerKey
{
    param([object]$ExchangeServer)

    if ($ExchangeServer.PSObject.Properties["DistinguishedName"] -and -not [string]::IsNullOrWhiteSpace("$($ExchangeServer.DistinguishedName)"))
    {
        return "$($ExchangeServer.DistinguishedName)".ToUpperInvariant()
    }

    return "$($ExchangeServer.Name)".ToUpperInvariant()
}

# ---------------------------------------------------------------------------
# Data collection and output
# ---------------------------------------------------------------------------
function Write-Authentication
{
    param([object]$Object)

    $Properties = [ordered]@{
        "Basic Authentication"              = @("BasicAuthentication","BasicAuthEnabled")
        "Windows Authentication"            = @("WindowsAuthentication","WindowsAuthEnabled")
        "Forms Authentication"              = @("FormsAuthentication")
        "Digest Authentication"             = @("DigestAuthentication")
        "OAuth Authentication"              = @("OAuthAuthentication")
        "ADFS Authentication"               = @("AdfsAuthentication")
        "WS-Security Authentication"        = @("WSSecurityAuthentication")
        "Certificate Authentication"        = @("CertificateAuthentication")
        "Client Certificate Authentication" = @("ClientCertAuth")
        "IIS Authentication Methods"        = @("IISAuthenticationMethods")
        "Internal Authentication Methods"   = @("InternalAuthenticationMethods")
        "External Authentication Methods"   = @("ExternalAuthenticationMethods")
    }

    $HeaderWritten = $false

    foreach ($Label in $Properties.Keys)
    {
        foreach ($Name in $Properties[$Label])
        {
            $Property = $Object.PSObject.Properties[$Name]

            if ($null -ne $Property)
            {
                if (-not $HeaderWritten)
                {
                    Write-Result " Authentication"
                    $HeaderWritten = $true
                }

                Write-Field $Label $Property.Value
                break
            }
        }
    }
}

# Reads only the ASA configuration text returned by Get-ClientAccessService.
# Returns the account name, "Configured", "Not Configured", or "Unknown".
# Credential health is not tested and passwords are never read.
function Get-ASAInfo
{
    param([object]$ClientAccessService)

    $Property = $ClientAccessService.PSObject.Properties["AlternateServiceAccountConfiguration"]

    if ($null -eq $Property)
    {
        return "Unknown"
    }

    if ($null -eq $Property.Value)
    {
        return "Not Configured"
    }

    $Configuration = "$($Property.Value)".Trim()

    if ([string]::IsNullOrWhiteSpace($Configuration) -or
        $Configuration -match '(?i)^<?Not\s*Set>?$' -or
        $Configuration -match '(?i)Latest:\s*<?Not\s*Set>?')
    {
        return "Not Configured"
    }

    if ($Configuration -match '(?is)Latest:\s*(?<Latest>.*?)(?:\s+Previous:|$)' -and
        $Matches["Latest"] -match '(?<Account>[^\s,]+\\[^\s,]+)\s*$')
    {
        return $Matches["Account"]
    }

    return "Configured"
}

function Get-VirtualDirectoryData
{
    param([string]$Command, [string]$ServerName)

    $Parameters = @{ Server = $ServerName; ErrorAction = "Stop" }

    if (-not $IncludeAuthentication)
    {
        $Parameters.AdPropertiesOnly = $true
    }

    return @(& $Command @Parameters)
}

function Write-StandardVirtualDirectory
{
    param([object[]]$Objects)

    if (-not $Objects -or $Objects.Count -eq 0)
    {
        Write-Field "Identity" "N/A"
        return
    }

    foreach ($Object in $Objects)
    {
        Write-Field "Identity" $Object.Identity
        Write-Field "Internal URL" $Object.InternalUrl
        Write-Field "External URL" $Object.ExternalUrl

        if ($IncludeAuthentication)
        {
            if ($null -ne $Object.PSObject.Properties["RequireSSL"])
            {
                Write-Field "Require SSL" $Object.RequireSSL
            }

            Write-Authentication $Object
        }
    }
}

function Write-OutlookAnywhere
{
    param([object[]]$Objects)

    if (-not $Objects -or $Objects.Count -eq 0)
    {
        Write-Field "Identity" "N/A"
        return
    }

    foreach ($Object in $Objects)
    {
        Write-Field "Identity" $Object.Identity
        Write-Field "Internal Hostname" $Object.InternalHostname
        Write-Field "External Hostname" $Object.ExternalHostname

        if ($IncludeAuthentication)
        {
            Write-Authentication $Object
            Write-Field "Internal Client Authentication" $Object.InternalClientAuthenticationMethod
            Write-Field "External Client Authentication" $Object.ExternalClientAuthenticationMethod
            Write-Field "Internal Clients Require SSL" $Object.InternalClientsRequireSsl
            Write-Field "External Clients Require SSL" $Object.ExternalClientsRequireSsl
            Write-Field "SSL Offloading" $Object.SSLOffloading
        }
    }
}

# Client Access Service: first with the ASA status, then without it. The
# Autodiscover virtual directory is queried separately in Write-ServiceData,
# so a failure here never hides it.
function Write-ClientAccessService
{
    param([string]$ServerName)

    $CAS = $null
    $ASA = "Unknown"
    $StatusError = $null
    $FallbackError = $null

    try
    {
        $CAS = Get-ClientAccessService -Identity $ServerName -IncludeAlternateServiceAccountCredentialStatus -ErrorAction Stop
        $ASA = Get-ASAInfo $CAS
    }
    catch
    {
        $StatusError = $_.Exception.Message

        try
        {
            $CAS = Get-ClientAccessService -Identity $ServerName -ErrorAction Stop
        }
        catch
        {
            $FallbackError = $_.Exception.Message
        }
    }

    Write-Result " Client Access Service"

    if ($null -ne $CAS)
    {
        Write-Field "Identity" $CAS.Identity
        Write-Field "AutoDiscover Service Internal URI" $CAS.AutoDiscoverServiceInternalUri
    }
    else
    {
        Write-Field "Identity" "Unknown"
        Write-Field "AutoDiscover Service Internal URI" "Unknown"
    }

    Write-Field "Alternate Service Account" $ASA

    if ($null -ne $FallbackError)
    {
        Write-ResultWarning "Client Access Service query failed on ${ServerName}: $FallbackError"
    }
    elseif ($null -ne $StatusError)
    {
        Write-ResultWarning "ASA status query failed on ${ServerName}: $StatusError"
    }
}

function Write-ServiceData
{
    param(
        [string]$ServiceName,
        [string]$ServerName
    )

    if ($ServiceName -eq "Autodiscover")
    {
        Write-ClientAccessService -ServerName $ServerName
        Write-Result ""
        Write-Result " Autodiscover Virtual Directory"
    }

    try
    {
        if ($ServiceName -eq "Autodiscover")
        {
            $Objects = Get-VirtualDirectoryData $ServiceConfig[$ServiceName].Command $ServerName

            if (-not $Objects -or $Objects.Count -eq 0)
            {
                Write-Field "Identity" "N/A"
            }
            else
            {
                foreach ($Object in $Objects)
                {
                    Write-Field "Identity" $Object.Identity

                    if ($IncludeAuthentication)
                    {
                        if ($null -ne $Object.PSObject.Properties["RequireSSL"])
                        {
                            Write-Field "Require SSL" $Object.RequireSSL
                        }

                        Write-Authentication $Object
                    }
                }
            }
        }
        elseif ($ServiceName -eq "OutlookAnywhere")
        {
            Write-OutlookAnywhere (Get-VirtualDirectoryData $ServiceConfig[$ServiceName].Command $ServerName)
        }
        else
        {
            Write-StandardVirtualDirectory (Get-VirtualDirectoryData $ServiceConfig[$ServiceName].Command $ServerName)
        }
    }
    catch
    {
        Write-ResultWarning "$ServiceName query failed on ${ServerName}: $($_.Exception.Message)"
    }
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
Show-StartupBanner

if ($PSBoundParameters.ContainsKey("OutputFile"))
{
    if ([string]::IsNullOrWhiteSpace($OutputFile))
    {
        Stop-Review "-OutputFile cannot be empty. Specify a file path."
    }

    try
    {
        $script:OutputTarget = Initialize-OutputFile -Path $OutputFile
    }
    catch
    {
        Stop-Review $_.Exception.Message
    }
}

# Load Exchange Management Shell if required.
if (-not (Get-Command Get-ExchangeServer -ErrorAction SilentlyContinue))
{
    $RemoteExchange = if ($env:ExchangeInstallPath)
    {
        Join-Path $env:ExchangeInstallPath "bin\RemoteExchange.ps1"
    }

    if (-not $RemoteExchange -or -not (Test-Path $RemoteExchange))
    {
        Stop-Review "Exchange Server management tools are not installed on this computer."
    }

    . $RemoteExchange
    Connect-ExchangeServer -Auto -AllowClobber
}

$ServiceConfig = [ordered]@{
    Autodiscover    = @{ Title = "Autodiscover";            Command = "Get-AutodiscoverVirtualDirectory" }
    ECP             = @{ Title = "Exchange Control Panel";  Command = "Get-EcpVirtualDirectory" }
    EWS             = @{ Title = "Exchange Web Services";   Command = "Get-WebServicesVirtualDirectory" }
    MAPI            = @{ Title = "MAPI";                    Command = "Get-MapiVirtualDirectory" }
    ActiveSync      = @{ Title = "ActiveSync";              Command = "Get-ActiveSyncVirtualDirectory" }
    OAB             = @{ Title = "Offline Address Book";    Command = "Get-OabVirtualDirectory" }
    OWA             = @{ Title = "Outlook on the web";      Command = "Get-OwaVirtualDirectory" }
    PowerShell      = @{ Title = "PowerShell";              Command = "Get-PowerShellVirtualDirectory" }
    OutlookAnywhere = @{ Title = "Outlook Anywhere";        Command = "Get-OutlookAnywhere" }
}

$SelectedServices = if ($Service)
{
    @($ServiceConfig.Keys | Where-Object { $Service -contains $_ })
}
else
{
    @($ServiceConfig.Keys)
}

# Resolve Exchange servers once. An explicit -Server list is never expanded
# to all servers, even when every entry is blank or invalid.
# Role filter kept from 2.1: ServerRole must contain Mailbox or ClientAccess.
$MailboxRolePattern = "Mailbox|ClientAccess"

if ($PSBoundParameters.ContainsKey("Server"))
{
    $ExchangeServers = New-Object System.Collections.ArrayList
    $SeenServers = @{}

    foreach ($ServerName in @($Server))
    {
        if ([string]::IsNullOrWhiteSpace($ServerName))
        {
            continue
        }

        $Resolved = Resolve-ExchangeServerName -Name $ServerName
        if ($null -eq $Resolved)
        {
            continue
        }

        $Key = Get-ServerKey $Resolved
        if ($SeenServers.ContainsKey($Key))
        {
            Write-Result "Note: Duplicate entry '$($ServerName.Trim())' is the same server as '$($SeenServers[$Key])' and was ignored." DarkGray
            continue
        }
        $SeenServers[$Key] = $ServerName.Trim()

        if ("$($Resolved.ServerRole)" -notmatch $MailboxRolePattern)
        {
            Write-ResultWarning "Exchange server '$($Resolved.Name)' was skipped. Detected role: $($Resolved.ServerRole). Only servers with the Mailbox or ClientAccess role are queried."
            continue
        }

        [void]$ExchangeServers.Add($Resolved)
    }

    $ExchangeServers = @($ExchangeServers)
}
else
{
    try
    {
        $ExchangeServers = @(Get-ExchangeServer -ErrorAction Stop | Sort-Object Name)
    }
    catch
    {
        Stop-Review "Exchange servers could not be read. $($_.Exception.Message)"
    }

    $ExchangeServers = @($ExchangeServers | Where-Object { "$($_.ServerRole)" -match $MailboxRolePattern })
}

if ($ExchangeServers.Count -eq 0)
{
    Stop-Review "No Exchange servers with the Mailbox or ClientAccess role were available to query."
}

Write-Field "Services" ($SelectedServices -join ", ")
Write-Field "Query Mode" $(if ($IncludeAuthentication) { "URLs + Authentication/IIS" } else { "Fast URL query (AD properties only)" })
if ($GroupByService)
{
    Write-Field "Group By" "Service"
}
Write-Result ""

# Console only, once per run. Shown after the OutputFile and server checks
# passed, right before the per-server Exchange queries start.
Write-Host 'Reading current Exchange values... This may take a while. Please wait.' -ForegroundColor DarkCyan
Write-Host ''

if ($GroupByService)
{
    foreach ($ServiceName in $SelectedServices)
    {
        Write-Result "======================================================================" Green
        Write-Result " Service: $($ServiceConfig[$ServiceName].Title)" Green
        Write-Result "======================================================================" Green
        Write-Result ""

        foreach ($ExchangeServer in $ExchangeServers)
        {
            $ServerName = $ExchangeServer.Name

            Write-Result " Server: $ServerName" Cyan
            Write-Field "FQDN" $ExchangeServer.Fqdn
            Write-ServiceData -ServiceName $ServiceName -ServerName $ServerName
            Write-Result ""
        }
    }
}
else
{
    foreach ($ExchangeServer in $ExchangeServers)
    {
        $ServerName = $ExchangeServer.Name

        Write-Result "======================================================================" Green
        Write-Result " Server: $ServerName" Green
        Write-Result "======================================================================" Green
        Write-Field "FQDN" $ExchangeServer.Fqdn
        Write-Field "Version" $ExchangeServer.AdminDisplayVersion
        Write-Field "Edition" $ExchangeServer.Edition
        Write-Result ""

        foreach ($ServiceName in $SelectedServices)
        {
            Write-Result $ServiceConfig[$ServiceName].Title Cyan
            Write-ServiceData -ServiceName $ServiceName -ServerName $ServerName
            Write-Result ""
        }
    }
}

# Result first, report save next, Feedback / Bugs footer last.
if ($script:WarningCount -eq 0)
{
    Write-Result "Review completed. All selected queries completed successfully." Green
}
else
{
    Write-Result ("Review completed with {0} warning(s). Some information could not be read; see the WARNING lines above." -f $script:WarningCount) Yellow
}
Write-Result "No Exchange configuration changes were made." DarkGray

$SaveFailed = $false
if ($null -ne $script:OutputTarget)
{
    try
    {
        Save-OutputFile -Path $script:OutputTarget
        Write-Host ""
        Write-Host "Output saved to: $($script:OutputTarget)" -ForegroundColor Yellow
    }
    catch
    {
        $SaveFailed = $true
        Write-Host ""
        Write-Host "ERROR: The report could not be saved to '$($script:OutputTarget)'. $($_.Exception.Message)" -ForegroundColor Red
    }
}

Show-CompletionFooter

if ($SaveFailed)
{
    exit 1
}
