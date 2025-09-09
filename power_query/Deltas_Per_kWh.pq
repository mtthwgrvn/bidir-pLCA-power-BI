let
  MergeAssump = Table.NestedJoin(PeakOff_ByYear, {"Scenario"}, Assumptions, {"Scenario"}, "A", JoinKind.LeftOuter),
  ExpandA = Table.ExpandTableColumn(MergeAssump, "A", {"η_V2X","η_BESS","ε"}, {"η_V2X","η_BESS","ε"}),

  Add_V1G = Table.AddColumn(ExpandA, "Δ_V1G_raw", each [Peak_end] - [Off_end], type number),
  Add_V1G_Adj = Table.AddColumn(Add_V1G, "Δ_V1G", each if [Peak_end] > [Off_end] + [ε] then [Δ_V1G_raw] else 0, type number),

  Add_V2H_raw = Table.AddColumn(Add_V1G_Adj, "Δ_V2H_raw", each [Peak_end] - ([Off_end]/[η_V2X]), type number),
  Add_V2H = Table.AddColumn(Add_V2H_raw, "Δ_V2H", each if [Peak_end] > [Off_end]/[η_V2X] + [ε] then [Δ_V2H_raw] else 0, type number),
  Add_V2B = Table.AddColumn(Add_V2H, "Δ_V2B", each [Δ_V2H], type number),

  Add_V2G_raw = Table.AddColumn(Add_V2B, "Δ_V2G_raw", each [Peak_end] * (1 - [Avg_Distloss]) - ([Off_end]/[η_V2X]), type number),
  Add_V2G = Table.AddColumn(Add_V2G_raw, "Δ_V2G", each if [Peak_end]*(1 - [Avg_Distloss]) > [Off_end]/[η_V2X] + [ε] then [Δ_V2G_raw] else 0, type number),

  Add_BESS_raw = Table.AddColumn(Add_V2G, "Δ_BESS_raw", each [Peak_bus] - ([Off_bus]/[η_BESS]), type number),
  Add_BESS = Table.AddColumn(Add_BESS_raw, "Δ_BESS", each if [Peak_bus] > [Off_bus]/[η_BESS] + [ε] then [Δ_BESS_raw] else 0, type number),

  Select = Table.SelectColumns(Add_BESS, {
    "Scenario","year","Peak_end","Off_end","Peak_bus","Off_bus","Avg_Distloss",
    "η_V2X","η_BESS",
    "Δ_V1G","Δ_V2H","Δ_V2B","Δ_V2G","Δ_BESS"
  })
in
  Select
