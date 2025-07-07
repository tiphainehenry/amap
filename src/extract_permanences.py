import pandas as pd
from pathlib import Path
from datetime import datetime, timedelta
import re

def main():

    def get_next_saturday():
        today = datetime.today()
        days_until_saturday = (5 - today.weekday()) % 7
        return (today + timedelta(days=days_until_saturday)).date()

    def extract_group(task_str):
        match = re.search(r"Distribution légumes ([a-zA-Z]+)", str(task_str), flags=re.IGNORECASE)
        return match.group(1).lower() if match else "unknown"

    def read_filtered_file(path: Path, target_date: datetime.date):
        try:
            df = pd.read_excel(path)
            df = df.dropna(how="all").dropna(axis=1, how="all")
            if "Date" not in df.columns or "Tâche" not in df.columns:
                print(f"⚠️ Skipping {path.name}: missing expected columns.")
                return None
            df["Date"] = pd.to_datetime(df["Date"], errors='coerce').dt.date
            df = df[df["Date"] == target_date]
            if df.empty:
                return None
            df["group"] = df["Tâche"].apply(extract_group)
            return df
        except Exception as e:
            print(f"❌ Failed to read {path.name}: {e}")
            return None

    def merge_amap_distributions(folder=".") -> pd.DataFrame:
        folder = Path(folder)
        files = list(folder.glob("Distribution_AMAP*.xlsx"))
        target_date = get_next_saturday()
        all_rows = [read_filtered_file(f, target_date) for f in files]
        all_rows = [df for df in all_rows if df is not None]
        merged_perms = pd.concat(all_rows, ignore_index=True) if all_rows else pd.DataFrame()
    
        output_path = folder / "merged_distributions_permanences.xlsx"
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            merged_perms.to_excel(writer, sheet_name="merged", index=False, header=True)

        print(f"✅ File saved: {output_path.name}")

    # Usage
    merge_amap_distributions(".")


if __name__ == "__main__":
    main()