#!/bin/bash

# Path to soffice (adjust if needed)
SOFFICE="/c/Program Files/LibreOffice/program/soffice.exe"

# Create archive folder if it doesn't exist
mkdir -p init

# Loop through all .xls files
for file in *.xls; do
  echo "Converting: $file"
  "$SOFFICE" --headless --convert-to xlsx:"Calc MS Excel 2007 XML" "$file"

  # If conversion succeeded, move original to init/
  if [ -f "${file%.xls}.xlsx" ]; then
    mv "$file" init/
  else
    echo "❌ Conversion failed for: $file"
  fi
done



# Run Python script after conversion
echo "Launching Python script..."
python src/main.py