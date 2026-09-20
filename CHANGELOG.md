# Changelog

## 2.1 - 2026-09-14

- Added Alternate Service Account (ASA) visibility to the Client Access Service output.
- Replaced the previous GroupBy modes with the simpler `-GroupByService` switch for comparing the same service configuration across multiple Exchange servers.
- Simplified the script structure and removed unnecessary processing layers.
- Improved default query performance by using `-AdPropertiesOnly` where applicable.
- Improved text-file output handling.
- Updated by Ceyhun Kirmizitas.

## 2.0 - 2026-08-25

- Updated `Get-ClientAccessServer` to `Get-ClientAccessService`.
- Added Autodiscover virtual directory details.
- Added optional authentication output with `-IncludeAuthentication`.
- Added optional text-file output with `-OutputFile`.
- Added N/A and Not Configured output.
- Made `-Server` optional. If omitted, all Exchange Mailbox servers in the organization are queried.
- Added `-Service` filtering.
- Added grouping and combined filtering support.
- Expanded built-in examples and parameter documentation.
- Updated by Ceyhun Kirmizitas.

## 1.10 - 2020-04-18

- Added PowerShell virtual directory and reordered output.
- Updated by Ali Tajran.

## 1.00 - 2015-08-27

- Initial version by Paul Cunningham.
