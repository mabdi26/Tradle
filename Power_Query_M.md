Power M Query - 

let
    // =========================
    // CONFIG
    // =========================
    BaseHost = "https://mfs3.aha.io",
    AhaToken = AhaApiToken,
    StartFY = 2026,
    EndFY   = 2028,

    // =========================
    // HELPERS
    // =========================

    // Robust custom field extractor by LABEL (case-insensitive)
    fnGetCFValue = (customFields as nullable list, fieldName as text) as nullable text =>
    let
        cfList = if customFields = null then {} else customFields,
        matched =
            List.First(
                List.Select(
                    cfList,
                    each try Text.Upper(Record.Field(_, "name")) = Text.Upper(fieldName) otherwise false
                ),
                null
            ),
        rawValue = if matched = null then null else try Record.Field(matched, "value") otherwise null,
        normalized =
            if rawValue = null then null
            else if Value.Is(rawValue, type text) then rawValue
            else if Value.Is(rawValue, type number) then Text.From(rawValue)
            else if Value.Is(rawValue, type record) then
                if Record.HasFields(rawValue, "name") then Text.From(rawValue[name]) else try Text.From(rawValue) otherwise null
            else if Value.Is(rawValue, type list) then
                Text.Combine(List.Transform(rawValue, each try Text.From(_) otherwise ""), ", ")
            else
                try Text.From(rawValue) otherwise null
    in
        normalized,

    // Clean common “null” / blanks
    fnCleanText = (x as any) as nullable text =>
    let
        t0 = if x = null then null else Text.Trim(Text.From(x)),
        t1 = if t0 = null or t0 = "" then null else t0,
        t2 = if t1 <> null and Text.Lower(t1) = "null" then null else t1
    in
        t2,

    // Normalize priority to H/M/L
    fnPriorityCode = (p as any) as nullable text =>
    let
        t = fnCleanText(p),
        s = if t = null then null else Text.Lower(t),
        out =
            if s = null then null
            else if s = "high" then "H"
            else if s = "medium" then "M"
            else if s = "low" then "L"
            else Text.Upper(Text.Start(t, 1))
    in
        out,

    // Status → C/D/P mapping - REVIEW W JULIE
    fnStatusToCDP = (statusText as any) as nullable text =>
    let
        t = fnCleanText(statusText),
        s = if t = null then "" else Text.Lower(t),
        code =
            if s = "" then null
            else if Text.Contains(s, "cancel") then null
            else if Text.Contains(s, "implement") then "C"
            else if Text.Contains(s, "ready") then "C"
            else if Text.Contains(s, "done") then null
            else if Text.Contains(s, "analyz") then "D"
            else if Text.Contains(s, "review") then "D"
            else if Text.Contains(s, "ideation") then "P"
            else "P"
    in
        code,

    // Robust date conversion
    fnToDate = (x as any) as nullable date =>
    let
        d1 = try Date.From(x) otherwise null,
        t  = try Text.From(x) otherwise null,
        d2 = if d1 <> null then d1 else (try Date.FromText(t, "en-US") otherwise null)
    in
        d2,

    // Safe field getter
    fnField = (r as nullable record, field as text) as any =>
        if r <> null and Record.HasFields(r, field) then Record.Field(r, field) else null,

    // Try a list of field names on a record and return first parsable date
    fnPickDateFromRecord = (r as nullable record, candidates as list) as nullable date =>
    let
        vals = List.Transform(candidates, each fnToDate(fnField(r, _))),
        hits = List.RemoveNulls(vals),
        out  = if List.Count(hits) > 0 then hits{0} else null
    in
        out,

    // Try nested record names + candidate fields inside them
    fnPickDateFromNested = (r as nullable record, nestedNames as list, candidates as list) as nullable date =>
    let
        nestedRecs =
            List.Transform(
                nestedNames,
                each
                    if r <> null and Record.HasFields(r, _) and Value.Is(Record.Field(r, _), type record)
                    then Record.Field(r, _)
                    else null
            ),
        nestedDates = List.Transform(nestedRecs, each fnPickDateFromRecord(_, candidates)),
        hits = List.RemoveNulls(nestedDates),
        out  = if List.Count(hits) > 0 then hits{0} else null
    in
        out,

    // Optional fallback: Parse 2 dates from initiative name text (MM/DD/YYYY) if present
    fnParseDateRangeFromText = (txt as nullable text) as record =>
    let
        s = if txt = null then "" else txt,
        cleaned = Text.Replace(Text.Replace(Text.Replace(Text.Replace(s, "—", " "), "-", " "), "–", " "), "|", " "),
        tokens = List.Select(Text.SplitAny(cleaned, " ,;()[]"), each _ <> null and _ <> ""),
        dateTokens = List.Select(tokens, each try fnToDate(_) <> null otherwise false),
        startD = if List.Count(dateTokens) >= 1 then fnToDate(dateTokens{0}) else null,
        endD   = if List.Count(dateTokens) >= 2 then fnToDate(dateTokens{1}) else null
    in
        [StartFromText = startD, EndFromText = endD],

    // Fetch full initiative details (detail endpoint)
    fnGetInitiativeDetail = (initiativeId as text) as nullable record =>
    let
        response =
            try Json.Document(
                Web.Contents(
                    BaseHost,
                    [
                        RelativePath = "api/v1/initiatives/" & initiativeId,
                        Headers = [ Authorization = "Bearer " & AhaToken, Accept = "application/json" ],
                        Timeout = #duration(0,0,2,0)
                    ]
                )
            )
            otherwise null,
        initiativeRecord =
            if response = null then null
            else if Record.HasFields(response, "initiative") then response[initiative]
            else null
    in
        initiativeRecord,

    // =========================
    // DIM QUARTERS 
    // =========================
    DimQuarters =
        let
            Years = List.Numbers(StartFY, EndFY - StartFY + 1),
            Quarters =
                List.TransformMany(
                    Years,
                    each {1..4},
                    (y, q) => [
                        Key = "Q.FY" & Text.End(Text.From(y), 2) & " Q" & Text.From(q),
                        Start = #date(y, (q - 1) * 3 + 1, 1),
                        End   = Date.EndOfMonth(#date(y, (q - 1) * 3 + 3, 1))
                    ]
                ),
            Tbl = Table.FromRecords(Quarters),
            Types = Table.TransformColumnTypes(Tbl, {{"Key", type text}, {"Start", type date}, {"End", type date}})
        in
            Types,

    QuarterKeys = DimQuarters[Key],

    fnRoadmapBars = (startDate as nullable date, endDate as nullable date, statusText as any) as record =>
    let
        code = fnStatusToCDP(statusText),
        s = startDate,
        e = endDate,
        Marks =
            if s = null or e = null or code = null then
                List.Repeat({null}, List.Count(QuarterKeys))
            else
                List.Transform(
                    Table.ToRecords(DimQuarters),
                    (r) => if s <= r[End] and e >= r[Start] then code else null
                ),
        Out = Record.FromList(Marks, QuarterKeys)
    in
        Out,

    // =========================
    // 1) PULL INITIATIVES (LIST)
    // =========================
    GetInitiativesPage = (PageNumber as number) =>
        let
            Source =
                Json.Document(
                    Web.Contents(
                        BaseHost,
                        [
                            RelativePath = "api/v1/initiatives",
                            Query = [ page = Text.From(PageNumber), per_page = "200" ],
                            Headers = [ Authorization = "Bearer " & AhaToken, Accept = "application/json" ],
                            Timeout = #duration(0,0,2,0)
                        ]
                    )
                )
        in
            Source[initiatives],

    Pages =
        List.Generate(
            () => [Page = 1, Data = GetInitiativesPage(1)],
            each List.Count([Data]) > 0,
            each [Page = [Page] + 1, Data = GetInitiativesPage([Page] + 1)],
            each [Data]
        ),

    Combined = List.Combine(Pages),
    T0 = Table.FromList(Combined, Splitter.SplitByNothing(), null, null, ExtraValues.Error),

    T1 =
        Table.ExpandRecordColumn(
            T0,
            "Column1",
            {"reference_num","name","start_date","due_date","workflow_status","custom_fields","url"},
            {"Initiative ID","Initiative name","Start (list)","End (list)","Workflow","CustomFields","AhaUrl"}
        ),

    T2 =
        Table.ExpandRecordColumn(
            T1,
            "Workflow",
            {"name"},
            {"Initiative status"}
        ),

    T3 =
        Table.TransformColumns(
            T2,
            {
                {"Start (list)", each fnToDate(_), type date},
                {"End (list)", each fnToDate(_), type date}
            }
        ),

    // taxonomy fields
    T4 = Table.AddColumn(T3, "MFS Category", each fnCleanText(fnGetCFValue([CustomFields], "MFS Category")), type text),
    T5 = Table.AddColumn(T4, "Supporting Objectives", each fnCleanText(fnGetCFValue([CustomFields], "Supporting Objectives")), type text),
    T6 = Table.AddColumn(T5, "Initiative Alignment", each fnCleanText(fnGetCFValue([CustomFields], "Initiative Alignment")), type text),
    T7 = Table.AddColumn(T6, "Priority", each fnPriorityCode(fnGetCFValue([CustomFields], "Priority")), type text),
    T8 = Table.AddColumn(T7, "OPS Initiative Rank", each fnCleanText(fnGetCFValue([CustomFields], "OPS Initiative Rank")), type text),
    T9 = Table.AddColumn(T8, "Impacted Domain(s)", each fnCleanText(fnGetCFValue([CustomFields], "Impacted Domain(s)")), type text),

    T10 =
        Table.AddColumn(
            T9,
            "OPS Initiative Rank (num)",
            each try Number.FromText([OPS Initiative Rank]) otherwise null,
            type number
        ),

    // =========================
    // 2) FILTER FIRST (Investment Operations)
    // =========================
    Filtered =
        Table.SelectRows(
            T10,
            each
                let d = try Text.Lower(Text.From([#"Impacted Domain(s)"])) otherwise ""
                in Text.Contains(d, "investment operations")
        ),

    // =========================
    // 3) FETCH FULL DETAILS (ACCURATE START/END)
    // =========================
    WithDetail =
        Table.AddColumn(
            Filtered,
            "Detail",
            each fnGetInitiativeDetail([Initiative ID]),
            type record
        ),

    // Robustly pull start/end from detail (tries many field names + nested objects)
    WithDetailDates =
        Table.AddColumn(
            Table.AddColumn(
                WithDetail,
                "Start (detail)",
                each
                    let
                        d = [Detail],
                        top = fnPickDateFromRecord(d, {"start_date","start_on","starts_on","begin_date","begin_on","start_at"}),
                        nested = fnPickDateFromNested(d, {"timeframe","date_range","schedule","timeline","timing"}, {"start_date","start_on","starts_on","begin_date","begin_on","start_at"})
                    in
                        if top <> null then top else nested,
                type date
            ),
            "End (detail)",
            each
                let
                    d = [Detail],
                    top = fnPickDateFromRecord(d, {"due_date","end_date","due_on","end_on","ends_on","finish_date","finish_on","end_at"}),
                    nested = fnPickDateFromNested(d, {"timeframe","date_range","schedule","timeline","timing"}, {"due_date","end_date","due_on","end_on","ends_on","finish_date","finish_on","end_at"})
                in
                    if top <> null then top else nested,
            type date
        ),

    // Optional fallback: parse from name if still missing (safe; does NOT fabricate)
    WithParsed =
        Table.AddColumn(
            WithDetailDates,
            "Parsed",
            each fnParseDateRangeFromText([Initiative name]),
            type record
        ),

    WithParsedDates =
        Table.ExpandRecordColumn(
            WithParsed,
            "Parsed",
            {"StartFromText","EndFromText"},
            {"Start (parsed)","End (parsed)"}
        ),

    // Choose best available dates (detail > list > parsed)
    WithFinalDates =
        Table.AddColumn(
            Table.AddColumn(
                WithParsedDates,
                "Initiative start date",
                each List.First(
                    List.RemoveNulls({
                        [#"Start (detail)"],
                        [#"Start (list)"],
                        [#"Start (parsed)"]
                    }),
                    null
                ),
                type date
            ),
            "Initiative end date",
            each List.First(
                List.RemoveNulls({
                    [#"End (detail)"],
                    [#"End (list)"],
                    [#"End (parsed)"]
                }),
                null
            ),
            type date
        ),

    // =========================
    // 4) CREATE FY/Q BARS
    // =========================
    WithBars =
        Table.AddColumn(
            WithFinalDates,
            "Bars",
            each fnRoadmapBars([Initiative start date], [Initiative end date], [Initiative status]),
            type record
        ),

    ExpandedBars =
        Table.ExpandRecordColumn(
            WithBars,
            "Bars",
            QuarterKeys,
            QuarterKeys
        ),

    BlankedFY =
        Table.ReplaceValue(
            ExpandedBars,
            null,
            "",
            Replacer.ReplaceValue,
            QuarterKeys
        ),

    Final =
        Table.SelectColumns(
            BlankedFY,
            List.Combine({
                {
                    "MFS Category",
                    "Supporting Objectives",
                    "Initiative Alignment",
                    "Priority",
                    "OPS Initiative Rank",
                    "OPS Initiative Rank (num)",
                    "Initiative name",
                    "Initiative start date",
                    "Initiative end date",
                    "Initiative status",
                    "Impacted Domain(s)",
                    "Initiative ID",
                    "AhaUrl"
                },
                QuarterKeys
            })
        ),
    #"Unpivoted Columns" = Table.UnpivotOtherColumns(Final, {"MFS Category", "Supporting Objectives", "Initiative Alignment", "Priority", "OPS Initiative Rank", "OPS Initiative Rank (num)", "Initiative name", "Initiative start date", "Initiative end date", "Initiative status", "Impacted Domain(s)", "Initiative ID", "AhaUrl"}, "Attribute", "Value"),
    #"Renamed Columns" = Table.RenameColumns(#"Unpivoted Columns",{{"Attribute", "Quarter"}, {"Value", "Status"}}),
    // =========================
//  FILTER: KEEP ONLY INITIATIVES WITH ANY C/D/P
// =========================

// 1) Determine if each initiative has ANY C/D/P across all quarters
HasCDPByInitiative =
    Table.Group(
        #"Renamed Columns",
        {"Initiative ID"},
        {
            {
                "HasCDP",
                each List.AnyTrue(
                    List.Transform(
                        [Status],
                        (s) => List.Contains({"C","D","P"}, s)
                    )
                ),
                type logical
            }
        }
    ),

// 2) Join this flag back onto all rows
WithHasCDP =
    Table.NestedJoin(
        #"Renamed Columns",
        {"Initiative ID"},
        HasCDPByInitiative,
        {"Initiative ID"},
        "HasCDPTable",
        JoinKind.LeftOuter
    ),

ExpandedHasCDP =
    Table.ExpandTableColumn(
        WithHasCDP,
        "HasCDPTable",
        {"HasCDP"},
        {"HasCDP"}
    ),

// 3) Keep ONLY initiatives that have at least one C/D/P
#"Keep Initiatives With Timeline" =
    Table.SelectRows(
        ExpandedHasCDP,
        each [HasCDP] = true
    ),

// 4) Cleanup helper column
#"Removed HasCDP" =
    Table.RemoveColumns(#"Keep Initiatives With Timeline", {"HasCDP"}),
    
    #"Filtered Rows" = Table.SelectRows(#"Renamed Columns", each ([Initiative status] <> "Cancelled" and [Initiative status] <> "Done") and ([MFS Category] <> null and [MFS Category] <> "Business Unit Priorities - Continuous Improvement" and [MFS Category] <> "Business Unit Priorities - Material Risk" and [MFS Category] <> "Business Unit Priorities - New Capability-Skill" and [MFS Category] <> "Non Discretionary - Production Support" and [MFS Category] <> "Non Discretionary - Run the Business")),
    #"Changed Type" = Table.TransformColumnTypes(#"Filtered Rows",{{"Status", type text}})
in
    #"Changed Type"
