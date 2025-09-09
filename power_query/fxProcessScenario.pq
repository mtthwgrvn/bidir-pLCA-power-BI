// fxProcessScenario(filePath as text, scenario as text) as table
(let
  Source = Csv.Document(File.Contents(filePath), [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
  Promote = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
  Types = Table.TransformColumnTypes(Promote, {
      {"timestamp_local", type datetime},
      {"year", Int64.Type},
      {"srmer_enduse_kg_per_MWh", type number},
      {"busbar_load_for_enduse_MW", type number},
      {"distloss_rate_marg", type number},
      {"srmer_busbar_kg_per_MWh", type number}
  }, "en-US"),

  Add_SRMER_kWh = Table.AddColumn(Types, "srmer_enduse_kg_per_kWh", each [srmer_enduse_kg_per_MWh] / 1000.0, type number),

  Add_Hour = Table.AddColumn(Add_SRMER_kWh, "hour", each Time.Hour(DateTime.Time([timestamp_local])), Int64.Type),
  Add_Dow = Table.AddColumn(Add_Hour, "dow", each Date.DayOfWeek(DateTime.Date([timestamp_local]), Day.Monday), Int64.Type),
  Add_Peak = Table.AddColumn(Add_Dow, "PeakOffFlag", each if ([dow] <= 4 and [hour] >= 15 and [hour] <= 20) then "Peak" else "Off-Peak", type text),

  Add_Busbar_kWh = Table.AddColumn(Add_Peak, "srmer_busbar_kg_per_kWh",
        each if Value.Is([srmer_busbar_kg_per_MWh], type number) and [srmer_busbar_kg_per_MWh] <> null
             then [srmer_busbar_kg_per_MWh] / 1000.0
             else [srmer_enduse_kg_per_kWh] * (1.0 - [distloss_rate_marg]),
        type number),

  Add_Weight_End = Table.AddColumn(Add_Busbar_kWh, "weighted_enduse", each [busbar_load_for_enduse_MW] * [srmer_enduse_kg_per_kWh], type number),
  Add_Weight_Bus = Table.AddColumn(Add_Weight_End, "weighted_busbar", each [busbar_load_for_enduse_MW] * [srmer_busbar_kg_per_kWh], type number),

  Add_Scenario = Table.AddColumn(Add_Weight_Bus, "Scenario", each scenario, type text),

  Keep = Table.SelectColumns(Add_Scenario, {
    "timestamp_local","year","Scenario","PeakOffFlag",
    "srmer_enduse_kg_per_kWh","srmer_busbar_kg_per_kWh",
    "busbar_load_for_enduse_MW","distloss_rate_marg",
    "weighted_enduse","weighted_busbar"
  })
 in
  Keep)
