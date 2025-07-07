import pandas as pd
import re
from datetime import datetime, timedelta
from pathlib import Path

def main():
    # Define cleaning and formatting function
    def clean_and_format(df: pd.DataFrame) -> pd.DataFrame:
        df = df.dropna(how="all").dropna(axis=1, how="all")

        # Detect header row
        header_row_idx = None
        for i in range(min(10, len(df))):
            row = df.iloc[i].fillna('').astype(str).str.lower()
            if "nom" in row.tolist() or "prénom" in row.tolist():
                header_row_idx = i
                break
        if header_row_idx is None:
            raise ValueError("No usable header row found (no 'Nom' or 'Prénom').")

        df.columns = df.iloc[header_row_idx].fillna('').astype(str).str.strip()
        df = df.iloc[header_row_idx + 1:].reset_index(drop=True)

        # Remove previous "Cumul" rows
        df = df[~df.iloc[:, 0].astype(str).str.lower().str.contains("cumul", na=False)]

        # Ensure column names are unique
        df.columns = df.columns.astype(str)
        if df.columns.duplicated().any():
            df.columns = [
                f"{col}_{i}" if df.columns.duplicated()[i] else col
                for i, col in enumerate(df.columns)
            ]

        def parse_number(val):
            if isinstance(val, str) and re.match(r"^\d{1,2}-\d{2}$", val.strip()):
                return float(val.strip().replace("-", "."))
            try:
                return float(val)
            except:
                return val

        def is_static_like(val):
            return isinstance(val, str) and re.match(r"^\d{1,2}-\d{2}$", val.strip())

        # Flag rows with any static-like string (e.g., "2-55")
        static_mask = df.map(is_static_like).any(axis=1)

        # Only parse numeric columns in non-static rows
        for col in df.columns[2:]:
            df.loc[~static_mask, col] = df.loc[~static_mask, col].map(parse_number)


        return df

    # --- Setup ---
    folder = Path('.')
    pattern = 'feuille-distribution-contrat-oeufs-2024-2025'
    today = datetime.today()
    days_until_saturday = (5 - today.weekday()) % 7
    saturday = today + timedelta(days=days_until_saturday)
    sheet_name = saturday.strftime("%Y-%m-%d")

    xlsx_files = list(folder.glob(f"{pattern}-*.xlsx"))

    merged_data = []
    cleaned_sheets = {}
    raw_sheets = {}
    static_lines = None

    # --- Load and process each file ---
    for file in xlsx_files:
        match = re.search(rf"{pattern}-([a-z0-9\- ]+)(?:\s\(\d+\))?\.xlsx", file.name, re.IGNORECASE)
        if not match:
            continue
        group = match.group(1).strip().lower()

        try:
            df = pd.read_excel(file, sheet_name=sheet_name, header=None)
        except ValueError:
            df = pd.read_excel(file, header=None)

        # --- Extract static header rows before the 'Nom'/'Prénom' line ---
        if static_lines is None:
            for i in range(len(df)):
                values = df.iloc[i].fillna('').astype(str).str.lower()
                if "cumul" in values.tolist():
                    static_lines = df.iloc[:i + 1].dropna(how="all").dropna(axis=1, how="all")
                    break

        df_clean = clean_and_format(df)
        df_clean["group"] = group
        merged_data.append(df_clean)

        cleaned_sheets[group] = df_clean.drop(columns="group", errors='ignore')
        raw_sheets[group] = df

    # --- Build merged output ---
    if merged_data:
        final_df = pd.concat(merged_data, ignore_index=True)

        # Ensure at least 7 columns so the summary can fit
        min_required_cols = 7
        if final_df.shape[1] < min_required_cols:
            extra_cols = [f"extra_{i}" for i in range(min_required_cols - final_df.shape[1])]
            final_df = final_df.reindex(columns=list(final_df.columns) + extra_cols)


        dedup_columns = [col for col in final_df.columns if col != "group"]
        final_df = final_df.drop_duplicates(subset=dedup_columns)

        # --- Detect and separate static-like rows (inner)
        def is_static_row(row):
            vals = row.fillna('').astype(str).str.lower().values
            return (
                all(v.strip() == "" for v in vals[:2])  # first columns (name fields) are empty
                and any(re.match(r"^\d{1,2}-\d{2}$", v) or "vrac" in v or "oeufs" in v for v in vals)
            )

        inner_static = final_df[final_df.apply(is_static_row, axis=1)]
        real_data = final_df[~final_df.apply(is_static_row, axis=1)]

        # Remove any row from real_data already present in static_lines
        if static_lines is not None:
            static_tuples = static_lines.apply(tuple, axis=1)
            real_data = real_data[~real_data.apply(tuple, axis=1).isin(static_tuples)]

        # --- Compute cumul only from rows where value == 1
        valid_rows = real_data[real_data.get("oeufs") == 1]
        cumul_counts = {
            col: (valid_rows[col] == 1).sum()
            for col in ["oeufs"]
            if col in valid_rows.columns
        }

        cumul_row = {col: "" for col in final_df.columns}
        cumul_row.update(cumul_counts)
        cumul_row[final_df.columns[0]] = "Cumul"
        cumul_row_df = pd.DataFrame([cumul_row])

        # Remove existing "Cumul" lines and replace with updated one
        static_lines = static_lines[~static_lines.iloc[:, 0].astype(str).str.lower().eq("cumul")]
        static_lines.columns = final_df.columns[:static_lines.shape[1]]

        # Insert one empty row before cumul and one before headers
        empty_row_df = pd.DataFrame([[""] * final_df.shape[1]], columns=final_df.columns)


    # Remove empty rows between cumul and real_data if already present
    real_data_cleaned = real_data.copy()

    # Clean top empty rows
    while not real_data_cleaned.empty and real_data_cleaned.iloc[0].isna().all():
        real_data_cleaned = real_data_cleaned.iloc[1:]


    # Compute summary
    n_boites_per_group = real_data_cleaned.groupby("group")["oeufs"].sum(min_count=1).fillna(0).astype(int)
    n_oeufs = n_boites_per_group * 6
    n_plaques = n_oeufs // 30
    n_restant = n_oeufs % 30

    # Format each row as a list of cells (with capitalized group names)
    summary_rows = []
    for group in n_boites_per_group.index:
        row = [
            group.upper(),
            f"{n_boites_per_group[group]}",
            "boites de 6",
            "=",
            f"{n_oeufs[group]} œufs, soit",
            f"{n_plaques[group]} plaques et",
            f"{n_restant[group]} œufs"
        ]
        summary_rows.append(row)

    n_cols = final_df.shape[1]

    # Safety check
    for row in summary_rows:
        if len(row) > n_cols:
            raise ValueError(f"Summary row too wide: {len(row)} columns vs only {n_cols} available.")

    # Pad all rows to match n_cols
    summary_df = pd.DataFrame(
        [row + [""] * (n_cols - len(row)) for row in summary_rows],
        columns=final_df.columns[:n_cols]
    )


    # Assemble final merged DataFrame
    full_merged = pd.concat([
        static_lines.reset_index(drop=True),
        empty_row_df,
        cumul_row_df,
        empty_row_df,
        summary_df,
        empty_row_df,
        real_data_cleaned.reset_index(drop=True), 
    ], ignore_index=True)


    # Deduplicate static rows again if needed
    def normalize_row_excluding_group(row, group_col="group"):
        return tuple(
            str(x).strip().lower()
            for i, x in enumerate(row)
            if pd.notna(x) and str(x).strip() and row.index[i] != group_col
        )

    reference_rows = set(normalize_row_excluding_group(row) for _, row in static_lines.iterrows())
    full_merged_dedup = full_merged[~full_merged.apply(
        lambda row: normalize_row_excluding_group(row) in reference_rows, axis=1
    )]
    full_merged_dedup = full_merged_dedup[
        ~full_merged_dedup.iloc[:, 0].astype(str).str.lower().eq("cumul")
    ]

    # Keep at most 1 empty row
    def is_empty_row(row):
        return all(str(val).strip() == "" for val in row)

    rows = full_merged_dedup.reset_index(drop=True)
    cleaned_rows = []
    prev_empty = False

    for i in range(len(rows)):
        is_empty = is_empty_row(rows.iloc[i])
        if is_empty and prev_empty:
            continue  # skip repeated empty rows
        cleaned_rows.append(rows.iloc[i])
        prev_empty = is_empty

    full_merged_dedup = pd.DataFrame(cleaned_rows, columns=rows.columns)

    full_merged = pd.concat([
        static_lines.reset_index(drop=True),
        empty_row_df,
        cumul_row_df,
        empty_row_df.copy(),
        full_merged_dedup.reset_index(drop=True)
    ], ignore_index=True)




    output_path = folder / "merged_distributions_oeufs.xlsx"
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        full_merged.to_excel(writer, sheet_name="merged", index=False, header=False)
        for group, raw_df in raw_sheets.items():
            if group in {"cscb", "four", "mjc"}:
                raw_sheet_name = f"{group}".replace(" ", "_")[:31]
                raw_df.to_excel(writer, sheet_name=raw_sheet_name, index=False)

    print(f"✅ File saved: {output_path.name}")


if __name__ == "__main__":
    main()