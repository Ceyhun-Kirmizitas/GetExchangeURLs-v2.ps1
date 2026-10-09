# GetExchangeURLs-v2.ps1

PowerShell script for reviewing Exchange Server client access URLs, authentication settings, and Alternate Service Account (ASA) configuration.

Intended for Exchange Server 2016, Exchange Server 2019, and Exchange Server Subscription Edition.

The script is read-only. It does not change Exchange configuration.

## Download

- [GitHub source](GetExchangeURLs-v2.ps1)
- [GitHub raw download](https://raw.githubusercontent.com/Ceyhun-Kirmizitas/GetExchangeURLs-v2.ps1/main/GetExchangeURLs-v2.ps1)

## What it shows

- Autodiscover
- ECP
- EWS
- MAPI
- ActiveSync
- OAB
- OWA
- PowerShell
- Outlook Anywhere
- Internal and external URLs or hostnames
- Optional authentication and IIS-backed settings
- Alternate Service Account configuration for Kerberos visibility

By default, the script queries all Exchange servers with the Mailbox or ClientAccess role. Use `-Server` to select specific servers. Invalid or unsupported servers are skipped with a warning, and duplicates are processed only once. Edge Transport is not queried.

By default, virtual directory queries use `-AdPropertiesOnly` where applicable for faster URL collection. Use `-IncludeAuthentication` when authentication and IIS-backed settings are required. Properties not exposed by a component are not shown.

The Alternate Service Account line shows the account name, `Not Configured`, or `Unknown` if ASA status could not be read. The script does not test ASA credential health.

Use `-OutputFile` to save results to a TXT file. The script checks the output path before querying Exchange and reports errors if saving fails. Existing files are overwritten.

## Parameters

| Parameter | Description |
|---|---|
| `-Server` | Selects one or more Exchange servers. All eligible servers are queried when omitted. |
| `-Service` | Selects one or more Client Access services to review. |
| `-GroupByService` | Groups results by service instead of server. |
| `-IncludeAuthentication` | Includes authentication and IIS-backed settings. |
| `-OutputFile` | Saves the results to a TXT file. |

## Examples

Review all eligible Exchange servers:

```powershell
.\GetExchangeURLs-v2.ps1
```

Review one server:

```powershell
.\GetExchangeURLs-v2.ps1 -Server EX01
```

Review selected services:

```powershell
.\GetExchangeURLs-v2.ps1 -Service MAPI,EWS
```

Compare the same services across multiple servers:

```powershell
.\GetExchangeURLs-v2.ps1 -Server EX01,EX02 -Service MAPI,EWS -GroupByService
```

Include authentication settings:

```powershell
.\GetExchangeURLs-v2.ps1 -Server EX01 -IncludeAuthentication
```

Save the review to a TXT file:

```powershell
.\GetExchangeURLs-v2.ps1 -Server EX01 -OutputFile C:\Temp\ExchangeURLs.txt
```

View the full PowerShell help:

```powershell
Get-Help .\GetExchangeURLs-v2.ps1 -Full
```

## Requirements

- Windows PowerShell 5.1
- Exchange Management Shell or available Exchange management cmdlets
- Required Exchange administrative permissions
- Exchange Server 2016, 2019, or Subscription Edition environment (intended compatibility; no hard version gate)

## Notes

Version 2.2 was live-lab validated on Exchange Server 2019 build 15.2.1748.10 (Mailbox role) using Windows PowerShell 5.1.

The following cases were validated with offline mock tests only:

- A configured Alternate Service Account (account name output)
- A failed ASA status query and a failed Client Access Service query
- An Edge Transport server or a wildcard name passed with `-Server`
- A report that cannot be saved at the end of the run

Exchange Server 2016 and Exchange Server Subscription Edition were not tested.

Always test the script in your environment before production use.

## Article

[Get Exchange Server URLs and Authentication Settings with PowerShell](https://ceyhunkirmizitas.net/get-exchange-server-urls-authentication-powershell/)

## Credits

Originally written by [Paul Cunningham](https://github.com/cunninghamp/ConfigureExchangeURLs.ps1).

Edited by Ali Tajran.

Further improved by Ceyhun Kirmizitas.

## Changelog

See [CHANGELOG.md](CHANGELOG.md).

## Feedback and issues

For bugs, feedback, or feature requests, use [GitHub Issues](https://github.com/Ceyhun-Kirmizitas/GetExchangeURLs-v2.ps1/issues).

Website: [ceyhunkirmizitas.net](https://ceyhunkirmizitas.net/)

## License

MIT. See [LICENSE](LICENSE).
