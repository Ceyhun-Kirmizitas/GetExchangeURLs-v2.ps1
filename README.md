# GetExchangeURLs-v2.ps1

PowerShell script for reviewing Exchange Server client access URLs, authentication settings, and Alternate Service Account (ASA) configuration.

Supports Exchange Server 2016, Exchange Server 2019, and Exchange Server Subscription Edition.

## Download

- [GitHub source](GetExchangeURLs-v2.ps1)
- [GitHub raw download](https://raw.githubusercontent.com/Ceyhun-Kirmizitas/GetExchangeURLs-v2.ps1/main/GetExchangeURLs-v2.ps1)
- [Website mirror](https://ceyhunkirmizitas.net/wp-content/uploads/tools/GetExchangeURLs-v2.ps1)

If GitHub access is restricted in your environment, the same script is also available from the website mirror.

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

If `-Server` is omitted, all Exchange Mailbox servers are queried.

By default, virtual directory queries use `-AdPropertiesOnly` where applicable for faster URL collection. Use `-IncludeAuthentication` when authentication and IIS-backed settings are also required.

## Examples

```powershell
.\GetExchangeURLs-v2.ps1
```

```powershell
.\GetExchangeURLs-v2.ps1 -Server EX01
```

```powershell
.\GetExchangeURLs-v2.ps1 -Service MAPI,EWS
```

```powershell
.\GetExchangeURLs-v2.ps1 -Server EX01,EX02 -GroupByService
```

```powershell
.\GetExchangeURLs-v2.ps1 -Server EX01 -IncludeAuthentication
```

```powershell
.\GetExchangeURLs-v2.ps1 -Server EX01 -OutputFile C:\Temp\ExchangeURLs.txt
```

## Article

[Get Exchange Server URLs and Authentication Settings with PowerShell](https://ceyhunkirmizitas.net/get-exchange-server-urls-authentication-powershell/)

## Credits

Originally written by [Paul Cunningham](https://github.com/cunninghamp/ConfigureExchangeURLs.ps1).

Edited by Ali Tajran.

Further improved by Ceyhun Kirmizitas.

See [CHANGELOG.md](CHANGELOG.md) for the changes in this version.

## License

MIT. See [LICENSE](LICENSE).
