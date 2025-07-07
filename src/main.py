import subprocess
import pandas as pd
from pathlib import Path
from datetime import datetime, timedelta

def run_extractors():
    print("Extraction des fichiers oeufs...")
    subprocess.run(["python", "src/extractor_oeufs.py"], check=True)

    print("Extraction des fichiers legumes...")
    subprocess.run(["python", "src/extractor_legumes.py"], check=True)

    print("Extraction de la liste des permanences...")
    subprocess.run(["python", "src/extract_permanences.py"], check=True)

def run_pdfs():
    print("Generating pdf...")
    subprocess.run(["python", "src/export_pdf.py"], check=True)


def get_next_saturday_sheetname():
    today = datetime.today()
    days_until_saturday = (5 - today.weekday()) % 7
    saturday = today + timedelta(days=days_until_saturday)
    return saturday.strftime("%Y-%m-%d")

def combine_outputs():
    merged_files = {
        "merged_distributions_permanences.xlsx": "permanences",
        "merged_distributions_oeufs.xlsx": "oeufs",
        "merged_distributions_legumes.xlsx": "legumes"
    }

    output_path = Path(f"distrib_amap_{get_next_saturday_sheetname()}.xlsx")

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for file, sheet_prefix in merged_files.items():
            path = Path(file)
            if not path.exists():
                print(f"⚠️ Missing file: {file}")
                continue

            # Use context manager to ensure ExcelFile is closed
            with pd.ExcelFile(path) as xls:
                for sheet_name in xls.sheet_names:
                    df = xls.parse(sheet_name, header=None)
                    sheet_base = sheet_name.replace("merged_", "")[:25]
                    
                    df.to_excel(writer, sheet_name=f"{sheet_prefix}_{sheet_base}", index=False, header=False)


    print(f"✅ Final file saved: {output_path.name}")

    # Now we can safely delete the files
    for f in merged_files:
        try:
            Path(f).unlink()
            # print(f"🗑️ Deleted: {f}")
        except Exception as e:
            print(f"❌ Could not delete {f}: {e}")

def main():
    run_extractors()
    combine_outputs()
    run_pdfs()

if __name__ == "__main__":
    main()
