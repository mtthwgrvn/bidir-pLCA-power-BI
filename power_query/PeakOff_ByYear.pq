let
  Combined = Table.Combine({Base, Optimistic, Pessimistic}),

  Grouped = Table.Group(
    Combined,
    {"Scenario","year","PeakOffFlag"},
    {
      {"Sum_Load_MW", each List.Sum([busbar_load_for_enduse_MW]), type number},
      {"Sum_Weighted_End", each List.Sum([weighted_enduse]), type number},
      {"Sum_Weighted_Bus", each List.Sum([weighted_busbar]), type number},
      {"Avg_Distloss", each List.Average([distloss_rate_marg]), type number}
    }
  ),

  Add_LWA_End = Table.AddColumn(Grouped, "LWA_enduse_kg_per_kWh",
      each if [Sum_Load_MW] = 0 then null else [Sum_Weighted_End] / [Sum_Load_MW], type number),
  Add_LWA_Bus = Table.AddColumn(Add_LWA_End, "LWA_busbar_kg_per_kWh",
      each if [Sum_Load_MW] = 0 then null else [Sum_Weighted_Bus] / [Sum_Load_MW], type number),

  PivotEnd = Table.Pivot(
      Table.TransformColumnTypes(
        Table.SelectColumns(Add_LWA_Bus, {"Scenario","year","PeakOffFlag","LWA_enduse_kg_per_kWh","LWA_busbar_kg_per_kWh","Avg_Distloss"}),
        {{"PeakOffFlag", type text}}
      ),
      List.Distinct(Add_LWA_Bus[PeakOffFlag]),
      "PeakOffFlag",
      "LWA_enduse_kg_per_kWh"
  ),

  PivotBus = Table.Pivot(
      Table.TransformColumnTypes(
        Table.SelectColumns(Add_LWA_Bus, {"Scenario","year","PeakOffFlag","LWA_busbar_kg_per_kWh"}),
        {{"PeakOffFlag", type text}}
      ),
      List.Distinct(Add_LWA_Bus[PeakOffFlag]),
      "PeakOffFlag",
      "LWA_busbar_kg_per_kWh"
  ),

  Join = Table.NestedJoin(PivotEnd, {"Scenario","year"}, PivotBus, {"Scenario","year"}, "Bus", JoinKind.LeftOuter),
  Expand = Table.ExpandTableColumn(Join, "Bus", {"Peak","Off-Peak"}, {"Peak_bus","Off_bus"}),

  Rename = Table.RenameColumns(Expand, {{"Peak","Peak_end"},{"Off-Peak","Off_end"}}),

  Final = Rename
in
  Final
