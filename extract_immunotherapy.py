import pandas as pd
import sys
from pathlib import Path

IMMUNOTHERAPY_DRUGS = [
    "Atezolizumab", "Pembrolizumab", "Durvalumab", "Nivolumab",
    "Relatlimab", "Cemiplimab", "Avelumab", "Ipilimumab"
]

def extract_immunotherapy(input_file: str, output_file: str = None, patient_file: str = None):
    input_path = Path(input_file)
    if output_file is None:
        output_file = input_path.stem + "_immunotherapy.xlsx"

    df = pd.read_excel(input_file)

    # Normalize column names (strip whitespace)
    df.columns = df.columns.str.strip()

    # Case-insensitive filter for any of the target drugs
    pattern = "|".join(IMMUNOTHERAPY_DRUGS)
    mask = df["EPIC_MEDICATION_NAME"].str.contains(pattern, case=False, na=False)
    filtered = df[mask].copy()

    # Sheet 1: selected columns, one row per order
    sheet1_cols = [
        "IP_PATIENT_ID", "ORDER_DATE", "START_DATE", "END_DATE",
        "EPIC_MED_ID", "EPIC_MEDICATION_NAME"
    ]
    # Keep only columns that exist
    sheet1_cols = [c for c in sheet1_cols if c in filtered.columns]
    sheet1 = filtered[sheet1_cols].reset_index(drop=True)

    # Sheet 2: group by patient + medication, earliest start / latest end
    group_cols = ["IP_PATIENT_ID", "EPIC_MED_ID", "EPIC_MEDICATION_NAME"]
    group_cols = [c for c in group_cols if c in filtered.columns]

    date_cols_present = {
        col: col in filtered.columns
        for col in ["START_DATE", "END_DATE"]
    }

    agg = {}
    if date_cols_present["START_DATE"]:
        agg["START_DATE"] = "min"
    if date_cols_present["END_DATE"]:
        agg["END_DATE"] = "max"

    sheet2 = (
        filtered.groupby(group_cols, as_index=False)
        .agg(agg)
        .sort_values(["IP_PATIENT_ID", "EPIC_MEDICATION_NAME"])
        .reset_index(drop=True)
    )

    if patient_file is not None:
        patient_df = pd.read_csv(patient_file, usecols=["IP_PATIENT_ID", "AGE", "SEX"])
        sheet2 = sheet2.merge(patient_df, on="IP_PATIENT_ID", how="left")
        # Move AGE and SEX to appear right after IP_PATIENT_ID
        cols = sheet2.columns.tolist()
        for col in ["SEX", "AGE"]:
            if col in cols:
                cols.insert(1, cols.pop(cols.index(col)))
        sheet2 = sheet2[cols]

    with pd.ExcelWriter(output_file, engine="openpyxl", datetime_format="YYYY-MM-DD") as writer:
        sheet1.to_excel(writer, sheet_name="Orders", index=False)
        sheet2.to_excel(writer, sheet_name="Summary", index=False)

    print(f"Input rows       : {len(df)}")
    print(f"Filtered rows    : {len(filtered)}")
    print(f"Unique patients  : {filtered['IP_PATIENT_ID'].nunique()}")
    print(f"Output written to: {output_file}")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python extract_immunotherapy.py <input.xlsx> [output.xlsx] [patients.csv]")
        sys.exit(1)
    input_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else None
    patient_file = sys.argv[3] if len(sys.argv) > 3 else None
    extract_immunotherapy(input_file, output_file, patient_file)
