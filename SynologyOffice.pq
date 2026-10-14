// Synology Office (Spreadsheet) connector for Power BI.
//
// Mirrors the Google Sheets experience: Get Data -> "Synology Office" -> paste the
// spreadsheet URL -> sign in -> Navigator lists the sheets of the .osheet file.
//
// Credentials are keyed by the NAS BASE URL (scheme://host[:port]), not by the full
// table link. So ONE sign-in / ONE gateway connection per NAS covers every table on
// it. The data source is SynologyOffice.Database(baseUrl, optional tableUrl): only
// the required baseUrl becomes the credential's data source path; the optional
// tableUrl is excluded from it. SynologyOffice.Contents(url) stays as a convenience
// wrapper so existing queries that pass a full table URL keep working unchanged.
//
// How it works (DSM Web API, same calls the Synology Drive web client uses):
//   1. SYNO.API.Auth            login                -> session id (sid)
//   2. SYNO.SynologyDrive.Files get                  -> resolve link / path to file_id
//   3. SYNO.Office.Export       download             -> .osheet converted to .xlsx
//   4. Excel.Workbook                                 -> navigation table of sheets
section SynologyOffice;

AppSession = "SynologyDrive";
RequestTimeout = #duration(0, 0, 5, 0);

// ---------------------------------------------------------------------------
// Public entry points
// ---------------------------------------------------------------------------

// Host-level data source. The only REQUIRED parameter is baseUrl, so Power Query
// keys the credential by the NAS base URL alone (optional parameters are not part
// of the data source path). Called with no tableUrl it just validates the sign-in
// (used by the gateway's TestConnection) and returns an empty navigation table.
[DataSource.Kind = "SynologyOffice", Publish = "SynologyOffice.Publish"]
shared SynologyOffice.Database = Value.ReplaceType(SynologyOfficeDatabaseImpl, SynologyOfficeDatabaseType);

// Convenience wrapper: take a full table URL, derive the NAS base from it and defer
// to the host-level data source. Existing queries keep calling this unchanged; the
// credential is still keyed by the base URL, because SynologyOffice.Database is the
// data source function invoked underneath.
shared SynologyOffice.Contents = (url as text) as table =>
    SynologyOffice.Database(BaseUrlOf(url), url);

SynologyOfficeDatabaseType = type function (
    baseUrl as (type text meta [
        Documentation.FieldCaption = Extension.LoadString("BaseUrlCaption"),
        Documentation.FieldDescription = Extension.LoadString("BaseUrlDescription"),
        Documentation.SampleValues = {"https://nas.example.com:5001"}
    ]),
    optional tableUrl as (type text meta [
        Documentation.FieldCaption = Extension.LoadString("TableUrlCaption"),
        Documentation.FieldDescription = Extension.LoadString("TableUrlDescription"),
        Documentation.SampleValues = {"https://nas.example.com:5001/oo/r/AbCdEf123456"}
    ])
) as table meta [
    Documentation.Name = Extension.LoadString("ButtonTitle"),
    Documentation.LongDescription = Extension.LoadString("LongDescription")
];

SynologyOfficeDatabaseImpl = (baseUrl as text, optional tableUrl as text) as table =>
    let
        normBase = NormalizeBase(baseUrl),
        sid = Login(normBase),
        result =
            if tableUrl = null then
                // No table requested: this is a plain credential check (gateway
                // TestConnection). M is lazy, so the sign-in must be forced explicitly:
                // "sid <> null" runs Login outside any try, so a wrong password still
                // fails here; only then is the session dropped (a failed logout is ignored).
                let
                    signedIn = sid <> null,
                    loggedOut = signedIn and ((try Logout(normBase, sid, #binary({})) otherwise false) or true)
                in
                    if loggedOut then EmptyNavigation() else error "unreachable"
            else
                let
                    parts = RequireHttps(Uri.Parts(tableUrl)),
                    fileRef = ResolveFileReference(parts),
                    info = GetFileInfo(normBase, sid, fileRef),
                    workbook = ExportXlsx(normBase, sid, info),
                    nav = ToNavigation(workbook)
                in
                    nav
    in
        result;

// Reduce a NAS address to scheme://host[:port], the exact string used as the
// credential's data source path. Trailing slash removed so RelativePath joins cleanly
// and so the key matches whether the base came from a full URL or was typed directly.
NormalizeBase = (baseUrl as text) as text =>
    let
        parts = RequireHttps(Uri.Parts(baseUrl)),
        host = parts[Host],
        withPort = parts[Scheme] & "://" & host
            & (if IsDefaultPort(parts[Scheme], parts[Port]) then "" else ":" & Text.From(parts[Port]))
    in
        withPort;

// scheme://host[:port] of a full table URL (used by the Contents wrapper).
BaseUrlOf = (url as text) as text => NormalizeBase(url);

IsDefaultPort = (scheme as text, port as number) as logical =>
    (Text.Lower(scheme) = "https" and port = 443) or (Text.Lower(scheme) = "http" and port = 80);

// The DSM password is posted in the request body, so plain http is allowed only for local testing.
RequireHttps = (parts as record) as record =>
    let
        host = Text.Lower(parts[Host]),
        isLocal = List.Contains({"localhost", "127.0.0.1", "[::1]", "::1"}, host) or Text.EndsWith(host, ".localhost")
    in
        if Text.Lower(parts[Scheme]) = "https" or isLocal then parts
        else error Error.Record("Expression.Error", Extension.LoadString("ErrorHttpsRequired"), [Scheme = parts[Scheme]]);

// Turns the URL the user pasted into a Drive "path" argument:
//   https://host/oo/r/<linkId>[#tid=1]   Office permanent link (address bar of an open sheet)
//   https://host/d/f/<linkId>            Drive permanent link
//   https://host/d/s/<linkId>/...        Drive sharing link
//   https://host/?file_id=123            numeric file id
//   https://host/?path=/mydrive/a.osheet path inside Drive
ResolveFileReference = (parts as record) as text =>
    let
        query = parts[Query],
        segments = List.Select(Text.Split(parts[Path], "/"), each _ <> ""),
        idxOf = (a as text, b as text) as number =>
            let
                pairs = List.Select(
                    List.Positions(segments),
                    (i) => i + 2 < List.Count(segments) and segments{i} = a and segments{i + 1} = b
                )
            in
                if List.IsEmpty(pairs) then -1 else pairs{0},
        oo = idxOf("oo", "r"),
        df = idxOf("d", "f"),
        ds = idxOf("d", "s"),
        linkAt = (i as number) as text => "link:" & segments{i + 2},
        result =
            if Record.HasFields(query, "file_id") then "id:" & query[file_id]
            else if Record.HasFields(query, "path") then query[path]
            else if oo >= 0 then linkAt(oo)
            else if df >= 0 then linkAt(df)
            else if ds >= 0 then linkAt(ds)
            else error Error.Record(
                "Expression.Error",
                Extension.LoadString("ErrorUnknownUrl"),
                [Url = parts[Path]]
            )
    in
        result;

// ---------------------------------------------------------------------------
// DSM Web API calls
// ---------------------------------------------------------------------------

Login = (baseUrl as text) as text =>
    let
        credential = Extension.CurrentCredential(),
        response = Web.Contents(
            baseUrl,
            [
                RelativePath = "webapi/entry.cgi",
                Headers = [#"Content-Type" = "application/x-www-form-urlencoded"],
                Content = Text.ToBinary(Uri.BuildQueryString([
                    api = "SYNO.API.Auth",
                    version = "6",
                    method = "login",
                    account = credential[Username],
                    passwd = credential[Password],
                    session = AppSession,
                    format = "sid"
                ])),
                // Every session is logged out after the export, so a cached login response
                // would hand out a dead sid; always ask DSM for a fresh one.
                IsRetry = true,
                ManualCredentials = true,
                ManualStatusHandling = {400, 401, 403, 404, 500, 502, 503},
                Timeout = RequestTimeout
            ]
        ),
        json = ParseDsmResponse(response, "SYNO.API.Auth"),
        sid = json[data][sid]
    in
        sid;

GetFileInfo = (baseUrl as text, sid as text, fileRef as text) as record =>
    let
        response = Web.Contents(
            baseUrl,
            [
                RelativePath = "webapi/entry.cgi",
                Headers = [#"Content-Type" = "application/x-www-form-urlencoded"],
                Content = Text.ToBinary(Uri.BuildQueryString([
                    api = "SYNO.SynologyDrive.Files",
                    version = "3",
                    method = "get",
                    path = fileRef,
                    _sid = sid
                ])),
                ManualCredentials = true,
                ManualStatusHandling = {400, 401, 403, 404, 500, 502, 503},
                Timeout = RequestTimeout
            ]
        ),
        json = ParseDsmResponse(response, "SYNO.SynologyDrive.Files"),
        data = json[data],
        name = data[name],
        checked =
            if Text.EndsWith(Text.Lower(name), ".osheet") then data
            else error Error.Record(
                "Expression.Error",
                Extension.LoadString("ErrorNotSpreadsheet"),
                [Name = name]
            )
    in
        checked;

// The sid goes in the POST body (not the URL) so it never lands in proxy or access logs.
ExportXlsx = (baseUrl as text, sid as text, info as record) as binary =>
    let
        exportName = Text.BeforeDelimiter(info[name], ".", {0, RelativePosition.FromEnd}) & ".xlsx",
        response = Web.Contents(
            baseUrl,
            [
                RelativePath = "webapi/entry.cgi/" & Uri.EscapeDataString(exportName),
                Headers = [#"Content-Type" = "application/x-www-form-urlencoded"],
                Content = Text.ToBinary(Uri.BuildQueryString([
                    api = "SYNO.Office.Export",
                    version = "1",
                    method = "download",
                    path = "id:" & Text.From(info[file_id]),
                    _sid = sid
                ])),
                ManualCredentials = true,
                ManualStatusHandling = {400, 401, 403, 404, 500, 502, 503},
                Timeout = RequestTimeout
            ]
        ),
        // Read the status before buffering: Binary.Buffer drops the response metadata.
        status = Value.Metadata(response)[Response.Status]? ?? 200,
        buffered = Binary.Buffer(response),
        // The session is no longer needed once the file is in memory. "or true" forces the
        // logout to run here, after the download, without letting its failure break the load.
        loggedOut = (try Logout(baseUrl, sid, buffered) otherwise false) or true,
        // A real xlsx is a zip archive and starts with "PK"; anything else is a DSM error JSON.
        isZip = Binary.Range(buffered, 0, 2) = #binary({0x50, 0x4B}),
        result =
            if not loggedOut then error "unreachable"
            else if status >= 400 then
                error Error.Record("DataSource.Error", "HTTP " & Text.From(status), [Api = "SYNO.Office.Export", File = info[name]])
            else if isZip then buffered
            else
                let parsed = try ParseDsmResponse(buffered, "SYNO.Office.Export") in
                    if parsed[HasError] then error parsed[Error]
                    else error Error.Record("DataSource.Error", Extension.LoadString("ErrorExport"), [File = info[name]])
    in
        result;

// "after" is only there to make the logout wait until the export has been downloaded.
Logout = (baseUrl as text, sid as text, after as binary) as logical =>
    let
        ready = Binary.Length(after) >= 0,
        response =
            if ready then
                Web.Contents(
                    baseUrl,
                    [
                        RelativePath = "webapi/entry.cgi",
                        Headers = [#"Content-Type" = "application/x-www-form-urlencoded"],
                        Content = Text.ToBinary(Uri.BuildQueryString([
                            api = "SYNO.API.Auth",
                            version = "6",
                            method = "logout",
                            session = AppSession,
                            _sid = sid
                        ])),
                        ManualCredentials = true,
                        ManualStatusHandling = {400, 401, 403, 404, 500, 502, 503},
                        Timeout = #duration(0, 0, 0, 30)
                    ]
                )
            else null,
        json = Json.Document(response)
    in
        json[success]? = true;

// Raises a readable error for HTTP failures and DSM {"success": false, "error": {"code": N}}.
ParseDsmResponse = (response as binary, api as text) as record =>
    let
        status = try Value.Metadata(response)[Response.Status] otherwise 200,
        json = try Json.Document(response) otherwise null,
        result =
            if status >= 400 then
                error Error.Record("DataSource.Error", "HTTP " & Text.From(status), [Api = api])
            else if json = null then
                error Error.Record("DataSource.Error", Extension.LoadString("ErrorNotDsm"), [Api = api])
            else if json[success]? = true then json
            else
                let code = json[error]?[code]? in
                    // Bad login makes Power BI show "Edit credentials" instead of a plain error.
                    if api = "SYNO.API.Auth" and List.Contains({400, 401, 403, 404, 406, 408, 410}, code) then
                        error Extension.CredentialError(Credential.AccessDenied, DsmErrorMessage(api, code))
                    else error Error.Record(
                        "DataSource.Error",
                        DsmErrorMessage(api, code),
                        [Api = api, Code = code]
                    )
    in
        result;

DsmErrorMessage = (api as text, code as nullable number) as text =>
    let
        // User-facing strings live in resources.resx (invariant/English) and
        // resources.ru-RU.resx (Russian); Extension.LoadString picks by locale.
        auth = [
            #"400" = Extension.LoadString("DsmAuth400"),
            #"401" = Extension.LoadString("DsmAuth401"),
            #"402" = Extension.LoadString("DsmAuth402"),
            #"403" = Extension.LoadString("DsmAuth403"),
            #"404" = Extension.LoadString("DsmAuth404"),
            #"406" = Extension.LoadString("DsmAuth406"),
            #"407" = Extension.LoadString("DsmAuth407"),
            #"408" = Extension.LoadString("DsmAuth408"),
            #"410" = Extension.LoadString("DsmAuth410")
        ],
        common = [
            #"102" = Extension.LoadString("DsmCommon102"),
            #"105" = Extension.LoadString("DsmCommon105"),
            #"106" = Extension.LoadString("DsmCommon106"),
            #"119" = Extension.LoadString("DsmCommon119")
        ],
        key = if code = null then "" else Text.From(code),
        codeText = if code = null then Extension.LoadString("DsmCodeUnknown") else key,
        hint = if api = "SYNO.SynologyDrive.Files" then Extension.LoadString("DsmFilesHint") else "",
        message =
            if api = "SYNO.API.Auth" and Record.HasFields(auth, key) then Record.Field(auth, key)
            else if Record.HasFields(common, key) then Record.Field(common, key)
            else Text.Format(Extension.LoadString("DsmErrorGeneric"), {api, codeText}) & hint
    in
        message;

// ---------------------------------------------------------------------------
// Navigation table (same shape as Google Sheets / Excel navigator)
// ---------------------------------------------------------------------------

ToNavigation = (xlsx as binary) as table =>
    let
        workbook = Excel.Workbook(xlsx, true, true),
        visible = Table.SelectRows(workbook, each not [Hidden]),
        trimmed = Table.SelectColumns(visible, {"Name", "Item", "Kind", "Data"}),
        withNav = Table.AddColumn(trimmed, "ItemKind", each [Kind], type text),
        withName = Table.AddColumn(withNav, "ItemName", each [ItemKind], type text),
        withLeaf = Table.AddColumn(withName, "IsLeaf", each true, type logical),
        nav = Table.ToNavigationTable(withLeaf, {"Item"}, "Name", "Data", "ItemKind", "ItemName", "IsLeaf")
    in
        nav;

// Empty navigator returned by a credentials-only call (no table requested). Same
// columns/shape as ToNavigation so the data source's result type stays consistent.
EmptyNavigation = () as table =>
    let
        empty = #table(
            type table [Name = text, Item = text, Kind = text, Data = any, ItemKind = text, ItemName = text, IsLeaf = logical],
            {}
        ),
        nav = Table.ToNavigationTable(empty, {"Item"}, "Name", "Data", "ItemKind", "ItemName", "IsLeaf")
    in
        nav;

Table.ToNavigationTable = (
    table as table,
    keyColumns as list,
    nameColumn as text,
    dataColumn as text,
    itemKindColumn as text,
    itemNameColumn as text,
    isLeafColumn as text
) as table =>
    let
        tableType = Value.Type(table),
        newTableType = Type.ReplaceTableKeys(tableType, {[Columns = keyColumns, Primary = true]}) meta [
            NavigationTable.NameColumn = nameColumn,
            NavigationTable.DataColumn = dataColumn,
            NavigationTable.ItemKindColumn = itemKindColumn,
            Preview.DelayColumn = itemNameColumn,
            NavigationTable.IsLeafColumn = isLeafColumn
        ],
        navigationTable = Value.ReplaceType(table, newTableType)
    in
        navigationTable;

// ---------------------------------------------------------------------------
// Data source definition
// ---------------------------------------------------------------------------

SynologyOffice = [
    // Lets the on-premises data gateway check credentials for scheduled refresh.
    // Keyed by the base URL, so one connection covers every table on the NAS.
    // dataSourcePath arrives as a JSON record string ({"baseUrl":"..."}) because the
    // data source function has named parameters, so parse out baseUrl before calling it.
    TestConnection = (dataSourcePath) =>
        let
            parsed = try Json.Document(dataSourcePath) otherwise null,
            baseUrl = if parsed <> null and Record.HasFields(parsed, "baseUrl") then parsed[baseUrl] else dataSourcePath
        in
            {"SynologyOffice.Database", baseUrl},
    Authentication = [
        UsernamePassword = [
            UsernameLabel = Extension.LoadString("UsernameLabel"),
            PasswordLabel = Extension.LoadString("PasswordLabel"),
            Label = Extension.LoadString("AuthLabel")
        ]
    ],
    Label = Extension.LoadString("DataSourceLabel")
];

SynologyOffice.Publish = [
    Beta = true,
    Category = "Online Services",
    ButtonText = {Extension.LoadString("ButtonTitle"), Extension.LoadString("ButtonHelp")},
    LearnMoreUrl = "https://kb.synology.com/en-global/DSM/help/Spreadsheet/office_general?version=7",
    SourceImage = SynologyOffice.Icons,
    SourceTypeImage = SynologyOffice.Icons
];

SynologyOffice.Icons = [
    Icon16 = {
        Extension.Contents("SynologyOffice16.png"),
        Extension.Contents("SynologyOffice20.png"),
        Extension.Contents("SynologyOffice24.png"),
        Extension.Contents("SynologyOffice32.png")
    },
    Icon32 = {
        Extension.Contents("SynologyOffice32.png"),
        Extension.Contents("SynologyOffice40.png"),
        Extension.Contents("SynologyOffice48.png"),
        Extension.Contents("SynologyOffice64.png")
    }
];
