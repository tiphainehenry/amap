#!/bin/bash

# Path to soffice (adjust if needed)
SOFFICE="/c/Program Files/LibreOffice/program/soffice.exe"

# Compute the date of next Saturday (in YYYY-MM-DD format)
NEXT_SATURDAY=$(date -d "next Saturday" +%Y-%m-%d)

# Create archive folder for next Saturday
mkdir -p "$NEXT_SATURDAY"

# Loop through all .xls files
for file in *.xls; do
  echo "Converting: $file"
  "$SOFFICE" --headless --convert-to xlsx:"Calc MS Excel 2007 XML" "$file"

  # If conversion succeeded, move original to the archive folder
  if [ -f "${file%.xls}.xlsx" ]; then
    mv "$file" "$NEXT_SATURDAY/"
  else
    echo "❌ Conversion failed for: $file"
  fi
done

# Run Python script after conversion
echo "Launching Python script..."
python src/main.py

# Move all generated .xlsx and .pdf files to the same archive folder
echo "Moving .xlsx and .pdf files to $NEXT_SATURDAY/"
mv *.xlsx *.pdf "$NEXT_SATURDAY/" 2>/dev/null
