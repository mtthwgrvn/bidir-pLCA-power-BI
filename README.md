# Emissions Excel Model

This repository contains a Python utility to build an Excel workbook from three scenario CSV files.

```
python build_workbook.py <MidCase.csv> <LowRECost_HighNGPrice.csv> <HighRECost_LowNGPrice.csv>
```

The script outputs `emissions_model.xlsx` with the following sheets:

- `raw_Base`, `raw_Optimistic`, `raw_Pessimistic`
- `combined`
- `agg_inputs`
- `peak_off_by_year`
- `Assumptions`
- `deltas_per_kWh`
- `Documentation`

Each sheet adheres to the requirements described in the project brief.
