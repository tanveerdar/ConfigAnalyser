This script takes an JSON based configuration export as input, analyses and extracts  configuration items from it , and creates a an excel sheet.

## Usage

python confignalyse.py -i {Config_file.tar.gz}

Or

python confignalyse.py -input {Config_file.tar.gz}

## Pyton Modules required

- openpyxl
- argparse
- json
- re
- tarfile