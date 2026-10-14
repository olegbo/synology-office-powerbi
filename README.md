# Synology Office — Power BI connector

A custom Power Query connector that loads **Synology Office (Spreadsheet, `.osheet`)** tables into Power BI, the same way the built-in Google Sheets connector works:

**Get data → Synology Office → NAS address → DSM sign-in → Navigator with the sheets.**

> Status: **working**. Verified end-to-end against a live DSM 7.x NAS — both in Power BI Desktop and via an on-premises data gateway (scheduled cloud refresh) — and against a DSM simulation (`tests/`). See [Notes](#notes-on-the-dsm-api).

> This is an independent, community project. It is **not affiliated with, endorsed by, or supported by Synology Inc.** "Synology", "DSM" and "Synology Office" are trademarks of Synology Inc., used here only to describe what the connector works with. It relies on DSM Web API calls that Synology does not document publicly, so a DSM or Synology Office update may break it.

The UI strings and error messages are localized: **English** by default, **Russian** when Power BI runs in a Russian locale (`resources.resx` / `resources.ru-RU.resx`).

## How it works

The connector uses the same DSM Web API calls as the Synology Drive web client:

| Step | API (`/webapi/entry.cgi`) | What it does |
|---|---|---|
| 1 | `SYNO.API.Auth` v6 `login` (POST, no cache) | DSM username/password → session id (`sid`) |
| 2 | `SYNO.SynologyDrive.Files` v3 `get` (POST) | resolve link/path → `file_id`, check it is an `.osheet` |
| 3 | `SYNO.Office.Export` v1 `download` (POST) | DSM converts the spreadsheet to `.xlsx` |
| 4 | `SYNO.API.Auth` v6 `logout` (POST) | session is closed right after the export; a failed logout does not fail the load |
| 5 | `Excel.Workbook` | navigation table of the visible sheets, first row as headers |

Every parameter, including `sid` and the password, is sent in the POST body rather than the URL, so it never lands in proxy logs. Connections are **HTTPS only** (`http://` is allowed only for `localhost`, used by the tests).

**Why not OAuth:** DSM has no OAuth for data access. The SSO Server (OIDC) only confirms identity and issues no tokens for the `SYNO.*` APIs. So the connector uses the Power BI **UsernamePassword** authentication kind.

**Why not the official Spreadsheet API:** it ships as a separate Docker container (`synology/spreadsheet-api`) whose documentation is behind a Synology account.

## Requirements

- DSM 7.x with the **Synology Drive Server** and **Synology Office** packages.
- The NAS reachable over **HTTPS only** (`http://` is rejected) **with a valid certificate** (e.g. Let's Encrypt in DSM); Power BI will not connect with a self-signed NAS certificate.
- A DSM user **without two-factor authentication (2FA)** that can open the tables. A dedicated read-only service account is recommended.
- Power BI Desktop (regular install or from the Microsoft Store).
- To **sign** the connector or **run the tests**: the Power Query SDK tools (`MakePQX.exe`, `PQTest.exe`). Get them once with:
  ```powershell
  powershell -ExecutionPolicy Bypass -File tools\get-sdk.ps1
  ```
  It downloads the NuGet package `Microsoft.PowerQuery.SdkTools` 2.155.2 (~130 MB) into `%LOCALAPPDATA%\PowerQuerySdkTools\2.155.2`, the folder `build.ps1` and the tests use by default. Building an unsigned `.mez` needs nothing extra.

## Credentials are keyed by the NAS base address

The data source is `SynologyOffice.Database(baseUrl, optional tableUrl)`: only the required `baseUrl` (`scheme://host[:port]`) becomes the credential's data source path — the optional table link is excluded. So **one sign-in / one gateway connection per NAS covers every table on it**. `SynologyOffice.Contents(fullUrl)` is a thin wrapper that derives the base from a full table URL, so existing queries keep working unchanged.

## Install

### Option A: unsigned `.mez` (simplest)

```powershell
powershell -File build.ps1 -Install
```

Then Power BI → Options → **GLOBAL → Security** → Data Extensions → "(Not Recommended) Allow any extension to load without validation or warning" → restart Power BI.

### Option B: signed `.pqx` (keeps the "Recommended" security level)

Only your own certificate is trusted; the security level stays at "Recommended".

1. Create a self-signed code-signing certificate (once):
   ```powershell
   powershell -File build.ps1 -NewCert
   ```
2. Build, sign and install:
   ```powershell
   powershell -File build.ps1 -Sign -Install
   ```
3. Once, in **PowerShell started "As Administrator"**, add the certificate's thumbprint to the Power BI policy (run it from the repo folder):
   ```powershell
   powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\trust-certificate.ps1
   ```
4. Power BI → Options → Security → Data Extensions → **(Recommended) … only certified and other trusted extensions**.
5. Fully close and reopen Power BI Desktop.

`build.ps1 -Install` copies the connector to `Documents\Power BI Desktop\Custom Connectors` and removes the other variant (`.mez`/`.pqx`) there so they do not clash. When signing, a temporary `.pfx` is exported from the certificate and deleted immediately after signing.

## Usage

1. **Get data → More…** → search `Synology` → **Synology Office** → Connect.
2. Enter the NAS base address (e.g. `https://nas:5001`). Optionally paste a table link to validate it, or leave it empty to only test the connection.
3. Enter your DSM username and password. The credential applies to every table on that NAS.
4. In a query, point at a table:
   ```m
   SynologyOffice.Contents("https://nas:5001/oo/r/AbCdEf123456")
   ```
   Then tick sheets in the Navigator → Load / Transform data.

### Supported links

| Format | Where to get it |
|---|---|
| `https://nas/oo/r/<id>[#tid=1]` | the address bar of an open spreadsheet |
| `https://nas/d/f/<id>` | a Synology Drive permalink |
| `https://nas/d/s/<id>/...` | a Drive share link (needs a signed-in user with access) |
| `https://nas/?file_id=<number>` | the numeric Drive file id |
| `https://nas/?path=/mydrive/Folder/Table.osheet` | a Drive path — fallback when a link is not recognized |

## Power BI Service (scheduled refresh via a gateway)

1. Install an **on-premises data gateway** that can reach the NAS.
2. Put the same connector on the gateway: copy `SynologyOffice.mez` (or the signed `.pqx`) into the gateway's custom-connectors folder, grant the gateway service account read access to it, and restart the gateway service. On the gateway machine, trust the certificate if you use the signed build.
3. In the Service → **Manage connections and gateways → New connection**: type **Synology Office**, **NAS address** = `https://nas:5001` (base address only, no path), authentication **DSM account**.
4. In the semantic model's **Gateway and cloud connections** settings, map the Synology data source to that connection and apply.

Because credentials are keyed by the base address, the connection's path must be the base URL. If a model was first published with a connector build that keyed by the full table URL, republish it once with this build so its declared data source path becomes the base URL; the gateway binding then matches.

## Where credentials are stored

The connector stores nothing and never caches the `sid` — every refresh signs in again.

- **Power BI Desktop:** in your Windows profile, encrypted with DPAPI. Regular install: `%LOCALAPPDATA%\Microsoft\Power BI Desktop\User.zip`; Store version: `%USERPROFILE%\Microsoft\Power BI Desktop Store App\User.zip`. The password never goes into the `.pbix`. Manage them under File → Options and settings → Data source settings.
- **Power BI Service:** the credentials are entered again on the on-premises gateway data source and stored encrypted with the gateway key.

## Limitations

- **Password-protected tables (Synology Office encryption) are not supported.** They are decrypted only in the browser; DSM's server-side export does not return them, and Power Query has no way to decrypt. Shared-folder/volume encryption in DSM works transparently and does not get in the way.
- **Anonymous access via a public link is not supported** — a DSM sign-in is required.
- **2FA** must be off for the connector account (a scheduled refresh cannot enter a code).
- Hidden sheets are not shown in the Navigator. Column types are not detected automatically — set them in the Power Query editor.
- Each load is a separate DSM sign-in and a full xlsx export of the spreadsheet. Selecting several sheets in the Navigator does this per sheet. To export a table once, create a single `SynologyOffice.Contents(...)` query (turn off "Enable load" for it) and get the sheets from it via Reference.

## Errors

| Message | What to do |
|---|---|
| Certificate error / "Could not establish trust" | Put a valid HTTPS certificate on the NAS |
| "Incorrect DSM username or password" | Check the credentials (Power BI will prompt to re-enter) |
| "…two-factor authentication (2FA)…" | Use an account without 2FA |
| "API not found" | Install Synology Drive Server and Synology Office |
| `SYNO.SynologyDrive.Files: DSM error, code …` | Check the link and permissions; try the `?path=/mydrive/…` form |
| "The link does not point to a spreadsheet…" | The link points to a document/slides, not an `.osheet` |
| "DSM did not return an xlsx…" | The table is probably password-protected |
| "Connections are only allowed over HTTPS…" | Use an `https://…` address |
| `HTTP 500` etc. | A DSM-side error during export — check the DSM / Synology Office logs |
| "The server did not return a DSM Web API response" | Wrong NAS address/port, or a reverse proxy with a path prefix |

DSM login error codes (400, 401, 403, 404, 406, 407, 408, 410) are turned into readable messages, and wrong-credential errors into `Extension.CredentialError` so Power BI shows "Edit credentials".

## Build

```powershell
powershell -File build.ps1                 # bin\SynologyOffice.mez
powershell -File build.ps1 -NewCert        # create the signing certificate (once)
powershell -File build.ps1 -Sign           # + bin\SynologyOffice.pqx (signed)
powershell -File build.ps1 -Sign -Install  # + install into Power BI Desktop
```

The `.mez` is a zip of `SynologyOffice.pq`, the `resources*.resx` files and the icons; the icons are generated on the first build. `-Sign` needs the Power Query SDK (see [Requirements](#requirements): `tools\get-sdk.ps1`); a different SDK folder can be passed with `-SdkTools <folder>`.

## Localization

User-facing strings live in resource files, loaded by `Extension.LoadString`:

- `resources.resx` — invariant / English (the fallback for any locale).
- `resources.ru-RU.resx` — Russian (used when Power BI runs in a Russian locale).

To add another language, copy `resources.resx` to `resources.<culture>.resx` (e.g. `resources.de-DE.resx`), translate the values, and rebuild — `build.ps1` packs every `resources*.resx`.

## Tests

```powershell
powershell -ExecutionPolicy Bypass -File tools\get-sdk.ps1      # once
powershell -ExecutionPolicy Bypass -File tests\run-tests.ps1
```

The script builds the connector, starts a DSM simulation (`tests/mock-dsm.ps1`, `http://localhost:8765`) with a test table (`tests/make-xlsx.ps1`: a visible sheet and a hidden one) and runs queries through **PQTest**: the navigator, sheet data, every link format, link errors, the HTTPS requirement, an HTTP export error, two loads in a row, a wrong password, 2FA, the base-address-only sign-in check (`SynologyOffice.Database(baseUrl)`) and `TestConnection` — both that it succeeds and that it **rejects a wrong password**. From the simulation log it checks that no `Authorization` header is sent, the `sid` never appears in a URL, and DSM is logged out after each export. The simulation issues a fresh `sid` per sign-in and invalidates it on logout, like real DSM. The exit code is the number of failed checks. Working files go to `tests/out/`.

The test credentials (`pbi` / `secret`) are fictitious and only stored in PQTest's credential store.

## Notes on the DSM API

The DSM simulation is modeled on the [synology-drive-api](https://github.com/zbjdonald/synology-drive-api) sources. The following were confirmed working against a live DSM 7.x NAS, but may vary by DSM version or setup:

1. `SYNO.SynologyDrive.Files get` accepts `path=link:<id>` for `/oo/r/<id>` links.
2. Signing in with `session=SynologyDrive` grants access to `SYNO.Office.Export`.
3. The xlsx export carries computed formula values (`Excel.Workbook` reads only stored values; otherwise formula cells would come back empty).
4. `SYNO.Office.Export download` accepts parameters in the POST body.
5. `/webapi` is reachable at the same host and port as the table link (including a separate Drive port or behind a reverse proxy).

## Project structure

```
SynologyOffice.pq       connector code (M)
resources.resx          strings — invariant / English
resources.ru-RU.resx    strings — Russian
SynologyOffice*.png     icons (generated by build.ps1)
build.ps1               build .mez / create cert / sign .pqx / install
trust-certificate.ps1   trust the certificate in the Power BI policy (as Administrator)
tests/                  DSM simulation, xlsx generator, run-tests.ps1
tools/get-sdk.ps1       download the Power Query SDK tools (for -Sign and tests)
```

## License

[MIT](LICENSE).
