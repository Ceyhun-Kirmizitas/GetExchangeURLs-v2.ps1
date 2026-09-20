<#
.SYNOPSIS
    Displays Exchange Client Access URLs and optional authentication settings.

.DESCRIPTION
    Displays Exchange Client Access URLs and namespaces for Exchange Server 2016,
    Exchange Server 2019, and Exchange Server Subscription Edition.

    If -Server is omitted, all Exchange Mailbox servers are queried.
    If -Service is omitted, all supported Client Access services are displayed.

    Default virtual directory queries use -AdPropertiesOnly for faster URL
    collection. Use -IncludeAuthentication when authentication and IIS-backed
    settings are also required.

    Client Access Service output includes Alternate Service Account (ASA)
    configuration status to help identify Kerberos deployments.

.PARAMETER Server
    Optional. Exchange Mailbox server(s) to query.

.PARAMETER Service
    Optional. Services to display: Autodiscover, ECP, EWS, MAPI, ActiveSync,
    OAB, OWA, PowerShell, OutlookAnywhere.

.PARAMETER GroupByService
    Optional. Groups selected results by service across Exchange servers.

.PARAMETER IncludeAuthentication
    Includes authentication and IIS-backed settings.

.PARAMETER OutputFile
    Saves the same formatted output to a text file.

.EXAMPLE
    .\GetExchangeURLs-v2.ps1

.EXAMPLE
    .\GetExchangeURLs-v2.ps1 -Server EX601

.EXAMPLE
    .\GetExchangeURLs-v2.ps1 -Service MAPI,EWS

.EXAMPLE
    .\GetExchangeURLs-v2.ps1 -Server EX601,EX602 -GroupByService

.EXAMPLE
    .\GetExchangeURLs-v2.ps1 -Server EX601,EX602 -Service MAPI,EWS -GroupByService

.EXAMPLE
    .\GetExchangeURLs-v2.ps1 -Server EX601 -IncludeAuthentication

.EXAMPLE
    .\GetExchangeURLs-v2.ps1 -Server EX601 -OutputFile C:\Temp\ExchangeURLs.txt

.NOTES
    Originally written by: Paul Cunningham
    Edited by: Ali Tajran
    Further updated by: Ceyhun Kirmizitas

    Change Log:

    V2.1, 14/09/2026 - Added Alternate Service Account (ASA) visibility to the
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

    V2, 25/08/2026 - Updated Get-ClientAccessServer to Get-ClientAccessService.
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

    V1.10, 18/04/2020 - Added PowerShell virtual directory and reordered output.
                         Updated by Ali Tajran.

    V1.00, 27/08/2015 - Initial version by Paul Cunningham.
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

# Load Exchange Management Shell if required.
if (-not (Get-Command Get-ExchangeServer -ErrorAction SilentlyContinue))
{
    $RemoteExchange = if ($env:ExchangeInstallPath)
    {
        Join-Path $env:ExchangeInstallPath "bin\RemoteExchange.ps1"
    }

    if (-not $RemoteExchange -or -not (Test-Path $RemoteExchange))
    {
        Write-Error "Exchange Server management tools are not installed on this computer."
        exit 1
    }

    . $RemoteExchange
    Connect-ExchangeServer -Auto -AllowClobber
}

# Output is buffered when -OutputFile is used, then written once at the end.
$script:OutputLines = New-Object 'System.Collections.Generic.List[string]'

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

function Get-ASAInfo
{
    param([object]$ClientAccessService)

    $Result = [PSCustomObject]@{ Account = "Not Configured"; Status = "Not Configured" }
    $Property = $ClientAccessService.PSObject.Properties["AlternateServiceAccountConfiguration"]

    if ($null -eq $Property -or $null -eq $Property.Value)
    {
        return $Result
    }

    $Configuration = "$($Property.Value)".Trim()

    if ([string]::IsNullOrWhiteSpace($Configuration) -or
        $Configuration -match '(?i)^<?Not\s*Set>?$' -or
        $Configuration -match '(?i)Latest:\s*<?Not\s*Set>?')
    {
        return $Result
    }

    $Result.Account = "Configured"
    $Result.Status = "Configured"

    if ($Configuration -match '(?is)Latest:\s*(?<Latest>.*?)(?:\s+Previous:|$)' -and
        $Matches["Latest"] -match '(?<Account>[^\s,]+\\[^\s,]+)\s*$')
    {
        $Result.Account = $Matches["Account"]
    }

    return $Result
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

function Write-ServiceData
{
    param(
        [string]$ServiceName,
        [string]$ServerName
    )

    try
    {
        if ($ServiceName -eq "Autodiscover")
        {
            $CAS = Get-ClientAccessService -Identity $ServerName -IncludeAlternateServiceAccountCredentialStatus -ErrorAction Stop
            $ASA = Get-ASAInfo $CAS

            Write-Result " Client Access Service"
            Write-Field "Identity" $CAS.Identity
            Write-Field "AutoDiscover Service Internal URI" $CAS.AutoDiscoverServiceInternalUri
            Write-Field "Alternate Service Account" $ASA.Account
            Write-Field "ASA Credential Status" $ASA.Status
            Write-Result ""
            Write-Result " Autodiscover Virtual Directory"

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
        Write-Result "WARNING: $ServiceName query failed on ${ServerName}: $($_.Exception.Message)" Yellow
    }
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

# Resolve Exchange servers once.
$ExchangeServers = if ($Server)
{
    foreach ($ServerName in $Server)
    {
        try
        {
            Get-ExchangeServer -Identity $ServerName -ErrorAction Stop
        }
        catch
        {
            Write-Result "WARNING: Exchange server '$ServerName' was not found." Yellow
        }
    }
}
else
{
    Get-ExchangeServer | Sort-Object Name
}

$ExchangeServers = @($ExchangeServers | Where-Object { "$($_.ServerRole)" -match "Mailbox|ClientAccess" })

if ($ExchangeServers.Count -eq 0)
{
    Write-Error "No Exchange Mailbox servers were available to query."
    exit 1
}

Write-Result ""
Write-Result "GetExchangeURLs-v2.ps1 - Version 2" Yellow
Write-Field "Services" ($SelectedServices -join ", ")
Write-Field "Query Mode" $(if ($IncludeAuthentication) { "URLs + Authentication/IIS" } else { "Fast URL query (AD properties only)" })
if ($GroupByService)
{
    Write-Field "Group By" "Service"
}
Write-Result ""

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

Write-Result "Finished querying all selected servers." Green

if ($OutputFile)
{
    $OutputDirectory = Split-Path -Path $OutputFile -Parent

    if ($OutputDirectory -and -not (Test-Path $OutputDirectory))
    {
        New-Item -Path $OutputDirectory -ItemType Directory -Force | Out-Null
    }

    Set-Content -Path $OutputFile -Value $script:OutputLines -Encoding UTF8
    Write-Host ""
    Write-Host "Output saved to: $OutputFile" -ForegroundColor Yellow
}
