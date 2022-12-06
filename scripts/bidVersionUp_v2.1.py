import os
import shutil
import sys
import warnings
from pathlib import Path
from typing import List

import openpyxl

warnings.simplefilter(
    action="ignore", category=UserWarning
)  # ignore warning about Data Validation


def main(filepath: str):
    # parse file, check file extension, and extract version number
    file = Path(filepath)
    if file.exists and file.suffix in [".xls", ".xlsm", ".xlsx"]:
        bid = file.name.split("_")
        print(f"Bid: {bid}")
    else:
        print("Not an excel file or not exists")
        sys.exit(1)

    # open bid in excel, update bid version + save versioned up file.
    bid_file = version_up(bid)
    # bid_file[0] - Bid File name without extension
    # bid_file[1] - Bid Version
    if os.path.isfile(filepath):
        new_bid = openpyxl.load_workbook(
            filename=filepath, read_only=False, keep_vba=True
        )  # updated line in v2.1 to include keep_vba
    else:
        print("File Not Found, check path and try again")
        sys.exit(1)  # Return an exit code

    breakdown = new_bid["Brkdwn"]

    breakdown.cell(row=4, column=10).value = "Bid_v" + str(bid_file[1])
    print("Updating Bid Version to: ", breakdown["J4"].value)
    print("Next Bid Version: Bid_v", bid_file[1])
    new_file = os.path.join(
        file.parent, f"{file.stem.split('_')[0]}_v{bid[1]}{file.suffix}"
    )
    new_bid.save(new_file)
    print(f"New Bid Located: {new_file}")

    # move previous version to old directory
    archive = os.path.join(file.parent, "archive")
    if not os.path.exists(archive):
        os.makedirs(archive)
    shutil.move(file, os.path.join(archive, file.name))
    sys.exit(0)  # Exit zero when success


def version_up(bid):
    version = 0
    version_up_bid = ""
    version = int(Path(bid[1]).stem.replace("v", "")) + 1
    version = str(version).zfill(3)  # zfill to add leading zeros
    bid[1] = version  # get version number from file name using bid position from loop
    version_up_bid = "_".join(bid)
    return [version_up_bid, version]


if __name__ == "__main__":  # Run a file directory like a script using this line
    # filepath = sys.argv[1] # Uncomment to run script from cmd
    filepath = "/home/adam/Downloads/test_v1.xlsx"  # Comment when using above
    main(filepath)
