# Emissions Excel Model (Power Query)

This repository provides Power Query (M) scripts to build a reproducible Excel workbook from three Cambium scenario CSVs. The workflow converts each file into a clean table, flags peak/off-peak hours, derives busbar intensities, and computes load-weighted averages and per-kWh deltas for V1G, V2X, and BESS programs.

## Contents
- `power_query/fxProcessScenario.pq` – function to import and transform a scenario CSV.
- `power_query/PeakOff_ByYear.pq` – query that combines scenarios and computes peak/off-peak load-weighted averages.
- `power_query/Deltas_Per_kWh.pq` – query that merges assumptions and calculates Managed/V2X/BESS deltas.
- `power_query/Assumptions.csv` – default efficiencies and tolerance values by scenario.

## Usage
1. In Excel Power Query, create a blank query and paste the contents of `fxProcessScenario.pq` to define the function.
2. For each scenario CSV, create another blank query that calls the function with the file path and scenario label:
   ```m
   let
     Base = fxProcessScenario("C:\\path\\to\\MidCase.csv", "Base")
   in
     Base
   ```
   Repeat for `Optimistic` and `Pessimistic` using the appropriate file names.
3. Import `PeakOff_ByYear.pq` and `Deltas_Per_kWh.pq` into new blank queries. Load `Assumptions.csv` as a table named `Assumptions`.
4. Enable **Load** on `Base`, `Optimistic`, `Pessimistic`, `PeakOff_ByYear`, and `Deltas_Per_kWh`. Refresh to regenerate results whenever the source CSVs are replaced.

`PeakOff_ByYear` outputs load-weighted peak/off-peak end-use and busbar intensities (kg CO₂e/kWh) along with average marginal distribution losses. `Deltas_Per_kWh` adds V1G, V2H/V2B, V2G, and BESS deltas with flip-tests to avoid negative values.

All intensities are reported in **kg CO₂e per kWh**. If time steps are not hourly, extend `fxProcessScenario` with a `Hours` column and weight by MW·h.
