import pandas as pd
from pathlib import Path
import argparse
from datetime import datetime

def load_scenario(path: Path, scenario: str) -> pd.DataFrame:
    df = pd.read_csv(path, parse_dates=["timestamp_local"])
    required = {
        "timestamp_local": "timestamp_local",
        "year": "year",
        "srmer_enduse_kg_per_MWh": "srmer_enduse_kg_per_MWh",
        "busbar_load_for_enduse_MW": "busbar_load_for_enduse_MW",
        "distloss_rate_marg": "distloss_rate_marg",
    }
    # rename columns if alternate names exist
    existing = {c.lower(): c for c in df.columns}
    for need, col in list(required.items()):
        if need not in existing:
            raise KeyError(f"Column '{need}' not found in {path.name}")
        required[need] = existing[need]
    df = df.rename(columns=required)
    # optional busbar srmer
    if "srmer_busbar_kg_per_MWh" in existing:
        df = df.rename(columns={existing["srmer_busbar_kg_per_MWh"]: "srmer_busbar_kg_per_MWh"})
    df["SRMER_enduse_kg_per_kWh"] = df["srmer_enduse_kg_per_MWh"] / 1000
    df["PeakOffFlag"] = df["timestamp_local"].apply(
        lambda t: "Peak" if (15 <= t.hour <= 20 and t.weekday() <= 4) else "Off-Peak"
    )
    if "srmer_busbar_kg_per_MWh" in df:
        df["Busbar_SRMER_kg_per_kWh"] = df["srmer_busbar_kg_per_MWh"] / 1000
    else:
        df["Busbar_SRMER_kg_per_kWh"] = df["SRMER_enduse_kg_per_kWh"] * (
            1 - df["distloss_rate_marg"]
        )
    df["Busbar_Weighting_CO2e"] = (
        df["busbar_load_for_enduse_MW"] * df["SRMER_enduse_kg_per_kWh"]
    )
    # hours column if irregular timestamp
    df = df.sort_values("timestamp_local")
    delta = df["timestamp_local"].diff().dt.total_seconds().div(3600).fillna(1)
    df["Hours"] = delta
    df["Scenario"] = scenario
    return df

def summarize(df: pd.DataFrame) -> pd.DataFrame:
    weight = df["busbar_load_for_enduse_MW"] * df["Hours"]
    df["Weight_kWh"] = weight
    group_cols = ["Scenario", "year", "PeakOffFlag"]
    agg = df.groupby(group_cols).apply(
        lambda g: pd.Series({
            "Sum_Load": g["Weight_kWh"].sum(),
            "Sum_Weighted_CO2e": (g["Busbar_Weighting_CO2e"] * g["Hours"]).sum(),
            "Sum_Busbar_CO2e": (g["Busbar_SRMER_kg_per_kWh"] * g["Weight_kWh"]).sum(),
            "Ldist": g["distloss_rate_marg"].mean(),
        })
    )
    agg["LWA_enduse_kg_per_kWh"] = agg["Sum_Weighted_CO2e"] / agg["Sum_Load"]
    agg["LWA_busbar_kg_per_kWh"] = agg["Sum_Busbar_CO2e"] / agg["Sum_Load"]
    return agg.reset_index()

def make_peak_off_table(agg: pd.DataFrame) -> pd.DataFrame:
    pivot = agg.pivot_table(
        index=["Scenario", "year"],
        columns="PeakOffFlag",
        values=["LWA_enduse_kg_per_kWh", "LWA_busbar_kg_per_kWh", "Ldist"],
    )
    pivot.columns = ["_".join(col).strip() for col in pivot.columns.values]
    pivot = pivot.reset_index()
    pivot = pivot.rename(
        columns={
            "LWA_enduse_kg_per_kWh_Peak": "Peak_end",
            "LWA_enduse_kg_per_kWh_Off-Peak": "Off_end",
            "LWA_busbar_kg_per_kWh_Peak": "Peak_bus",
            "LWA_busbar_kg_per_kWh_Off-Peak": "Off_bus",
            "Ldist_Peak": "Ldist_peak",
        }
    )
    return pivot

def add_deltas(summary: pd.DataFrame, assumptions: pd.DataFrame) -> pd.DataFrame:
    df = summary.merge(assumptions, on="Scenario", how="left")
    eps = df["ε_tolerance"]
    df["Managed"] = (
        (df["Peak_end"] > df["Off_end"] + eps) * (df["Peak_end"] - df["Off_end"])
    )
    df["V2H"] = (
        (df["Peak_end"] > df["Off_end"] / df["η_V2X"] + eps)
        * (df["Peak_end"] - df["Off_end"] / df["η_V2X"])
    )
    df["V2B"] = df["V2H"]
    df["V2G"] = (
        (
            df["Peak_end"] * (1 - df["Ldist_peak"]) >
            df["Off_end"] / df["η_V2X"] + eps
        )
        * (df["Peak_end"] * (1 - df["Ldist_peak"]) - df["Off_end"] / df["η_V2X"])
    )
    df["BESS"] = (
        (
            df["Peak_bus"] > df["Off_bus"] / df["η_BESS"] + eps
        )
        * (df["Peak_bus"] - df["Off_bus"] / df["η_BESS"])
    )
    return df

def build_workbook(base_csv: Path, optimistic_csv: Path, pessimistic_csv: Path, output: Path):
    scenarios = {
        "Base": base_csv,
        "Optimistic": optimistic_csv,
        "Pessimistic": pessimistic_csv,
    }
    raw_frames = []
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for name, path in scenarios.items():
            df = load_scenario(path, name)
            raw_frames.append(df)
            df.to_excel(writer, sheet_name=f"raw_{name}", index=False)
        combined = pd.concat(raw_frames, ignore_index=True)
        combined.to_excel(writer, sheet_name="combined", index=False)
        agg = summarize(combined)
        agg.to_excel(writer, sheet_name="agg_inputs", index=False)
        summary = make_peak_off_table(agg)
        summary.to_excel(writer, sheet_name="peak_off_by_year", index=False)
        assumptions = pd.DataFrame(
            {
                "Scenario": ["Base", "Optimistic", "Pessimistic"],
                "η_V2X": [0.90, 0.92, 0.85],
                "η_BESS": [0.88, 0.91, 0.85],
                "ε_tolerance": [0.001, 0.001, 0.001],
            }
        )
        assumptions.to_excel(writer, sheet_name="Assumptions", index=False)
        deltas = add_deltas(summary, assumptions)
        deltas.to_excel(writer, sheet_name="deltas_per_kWh", index=False)
        doc = pd.DataFrame(
            {
                "Version": ["1.0"],
                "Generated": [datetime.utcnow().isoformat()],
                "Base_csv": [base_csv.name],
                "Optimistic_csv": [optimistic_csv.name],
                "Pessimistic_csv": [pessimistic_csv.name],
            }
        )
        doc.to_excel(writer, sheet_name="Documentation", index=False)

def main():
    p = argparse.ArgumentParser(description="Build emissions Excel model")
    p.add_argument("base_csv", type=Path)
    p.add_argument("optimistic_csv", type=Path)
    p.add_argument("pessimistic_csv", type=Path)
    p.add_argument("--output", type=Path, default=Path("emissions_model.xlsx"))
    args = p.parse_args()
    build_workbook(args.base_csv, args.optimistic_csv, args.pessimistic_csv, args.output)

if __name__ == "__main__":
    main()
