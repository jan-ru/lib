"""
Check that file, sheet, table exists
"""

import sys
import magic

def file_exists(filename: str) -> bool:
    """credit: excel2markdown by SpeerSec:
    https://github.com/SpeerSec/excel2markdown
    """

    # Check the file extension
    if not filename.endswith((".xlsx", ".xml")):
        sys.exit(
            'Error: only Excel files with the .xlsx or .xml extension are allowed.'
        )

    # Open the file in binary mode
    with open(filename, "rb") as t:
        file_type = magic.from_buffer(t.read())

    # Check if the file is an Excel file
    if "Microsoft Excel" not in file_type and "XML" not in file_type:
        sys.exit('Error: the provided file is not an Excel file.')

    # mss naar logfile schrijven ipv naar console
    # print("file exists")
    return True