import os
import csv
import logging
import pandas as pd
import warnings
from datetime import datetime
from rich import print


def fxn():
    warnings.warn("deprecated", DeprecationWarning)
    warnings.warn("future", FutureWarning)


with warnings.catch_warnings():
    warnings.simplefilter("ignore")
    fxn()


LOG_FP = r"\\busse\Purchasing\Physical Inventory\logs"

logging.basicConfig(
    filename=os.path.join(LOG_FP, f"metrics_{datetime.now():%Y-%m-%d_%H-%M-%S}.log"),
    level=logging.INFO,
    format="%(asctime)s \n%(message)s",
)

LOGGER = logging.getLogger(__name__)
FILE_NAME = "PICOUNT"
FILE_PATH = os.path.join(r"c:\temp", FILE_NAME)

if not os.path.exists(FILE_PATH):
    print("File not found.")
    print(f"Path: {FILE_PATH}")
    print(f"Expected file name: {FILE_NAME}")
    print("Run PI.DISCREP.RPT to generate file.")
    print("Exiting...")
    exit()


def main():

    header = [
        "part",
        "description",
        "uom",
        "chg.value",
        "pre-count.value",
        "post-count.value",
        "diff.value",
        "pre-count.qty",
        "post-count.qty",
        "diff.qty",
    ]

    rows = []

    with open(FILE_PATH, "r") as f:
        reader = csv.reader(f)

        for row in reader:
            rows.append(dict(zip(header, row)))

    pre_count_value = 0
    post_count_value = 0
    diff_value = 0
    pre_count_qty = 0
    post_count_qty = 0
    diff_qty = 0

    for row in rows:
        row["chg.value"] = float(row["chg.value"])
        row["pre-count.value"] = float(row["pre-count.value"])
        row["post-count.value"] = float(row["post-count.value"])
        row["diff.value"] = float(row["diff.value"])
        row["pre-count.qty"] = int(row["pre-count.qty"])
        row["post-count.qty"] = int(row["post-count.qty"])
        row["diff.qty"] = int(row["diff.qty"])

        pre_count_value += row["pre-count.value"]
        post_count_value += row["post-count.value"]
        diff_value += row["diff.value"]
        pre_count_qty += row["pre-count.qty"]
        post_count_qty += row["post-count.qty"]
        diff_qty += row["diff.qty"]

    # rows.append(["", "", "", "", "", "", "", "", "", ""])
    # rows.append(
    #     [
    #         "Total",
    #         "",
    #         "",
    #         "",
    #         pre_count_value,
    #         post_count_value,
    #         diff_value,
    #         pre_count_qty,
    #         post_count_qty,
    #         diff_qty,
    #     ]
    # )

    df = pd.DataFrame(rows)

    df = df.append(
        pd.Series(
            [
                "Total",
                "",
                "",
                "",
                pre_count_value,
                post_count_value,
                diff_value,
                pre_count_qty,
                post_count_qty,
                diff_qty,
            ],
            index=df.columns,
        ),
        ignore_index=True,
    )

    return df


def print_metrics(df: pd.DataFrame) -> None:
    print("---------------------------------")
    total_pre_count_value = df["pre-count.value"].sum() - df["pre-count.value"].iloc[-1]
    print(f"Total Pre Count Value: {total_pre_count_value}")
    LOGGER.info(f"Total Pre Count Value: {total_pre_count_value}")
    total_post_count_value = (
        df["post-count.value"].sum() - df["post-count.value"].iloc[-1]
    )
    print(f"Total Post Count Value: {total_post_count_value}")
    LOGGER.info(f"Total Post Count Value: {total_post_count_value}")
    print("---------------------------------")
    diff_sum = df["diff.value"].sum() - df["diff.value"].iloc[-1]
    print(f"Total Diff: {diff_sum}")
    LOGGER.info(f"Total Diff: {diff_sum}")
    print("---------------------------------")
    percent_diff = (diff_sum / total_pre_count_value) * 100
    print(f"Percent Diff: {percent_diff}%")
    LOGGER.info(f"Percent Diff: {percent_diff}%")


def save_to_excel(
    df: list[pd.DataFrame],
    file_name: str,
    sheet_names: list[str],
    year: str = input("What year is it?\t > "),
) -> None:
    base_path = r"c:\temp"

    if os.path.exists(rf"\\busse\purchasing\Physical Inventory\{year}\discrepancies"):
        base_path = rf"\\busse\purchasing\Physical Inventory\{year}\discrepancies"
        print("Using network path for output file.")
        print(f"Path: {base_path}")

    save_path = os.path.join(
        base_path, f"{file_name}_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
    )

    with pd.ExcelWriter(save_path, engine="xlsxwriter") as writer:
        for i, df in enumerate(df):
            try:
                df.to_excel(writer, sheet_name=sheet_names[i], index=False)
            except KeyError:
                df.to_excel(writer, sheet_name=f"Sheet{i+1}", index=False)


if __name__ == "__main__":
    df_all = main()
    print_metrics(df_all)

    save_to_excel([df_all], FILE_NAME, ["All"])
