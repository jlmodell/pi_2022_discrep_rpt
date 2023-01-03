import os
import csv
import logging
import pandas as pd
from datetime import datetime
from rich import print

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
    rows_non_zero = []

    with open(FILE_PATH, "r") as f:
        reader = csv.reader(f)

        for row in reader:
            if int(row[7]) != 0:
                rows_non_zero.append(dict(zip(header, row)))
            rows.append(dict(zip(header, row)))

    for row in rows:
        post_count_qty = 0
        diff_qty = 0
        if float(row["chg.value"]) > 0:
            post_count_qty = int(
                float(row["post-count.value"]) / float(row["chg.value"])
            )
            diff_qty = post_count_qty - int(row["pre-count.qty"])

        row["chg.value"] = float(row["chg.value"])
        row["pre-count.value"] = float(row["pre-count.value"])
        row["post-count.value"] = float(row["post-count.value"])
        row["diff.value"] = float(row["diff.value"])
        row["pre-count.qty"] = int(row["pre-count.qty"])
        row["post-count.qty"] = post_count_qty
        row["diff.qty"] = diff_qty

    for row in rows_non_zero:
        post_count_qty = 0
        diff_qty = 0
        if float(row["chg.value"]) > 0:
            post_count_qty = int(
                float(row["post-count.value"]) / float(row["chg.value"])
            )
            diff_qty = post_count_qty - int(row["pre-count.qty"])

        row["chg.value"] = float(row["chg.value"])
        row["pre-count.value"] = float(row["pre-count.value"])
        row["post-count.value"] = float(row["post-count.value"])
        row["diff.value"] = float(row["diff.value"])
        row["pre-count.qty"] = int(row["pre-count.qty"])
        row["post-count.qty"] = post_count_qty
        row["diff.qty"] = diff_qty

    return rows, rows_non_zero


def sort_largest_value(df: pd.DataFrame) -> pd.DataFrame:
    return df.sort_values(by="diff.value", ascending=False).copy()


def sort_smallest_value(df: pd.DataFrame) -> pd.DataFrame:
    return df.sort_values(by="diff.value", ascending=True).copy()


def print_metrics(df: pd.DataFrame) -> None:
    print("---------------------------------")
    total_pre_count_value = df["pre-count.value"].sum()
    print(f"Total Pre Count Value: {total_pre_count_value}")
    LOGGER.info(f"Total Pre Count Value: {total_pre_count_value}")
    total_post_count_value = df["post-count.value"].sum()
    print(f"Total Post Count Value: {total_post_count_value}")
    LOGGER.info(f"Total Post Count Value: {total_post_count_value}")
    print("---------------------------------")
    diff_sum = df["diff.value"].sum()
    print(f"Total Diff: {diff_sum}")
    LOGGER.info(f"Total Diff: {diff_sum}")
    print("---------------------------------")
    percent_diff = (diff_sum / total_pre_count_value) * 100
    print(f"Percent Diff: {percent_diff}%")
    LOGGER.info(f"Percent Diff: {percent_diff}%")


def print_top(df: pd.DataFrame, descriptor: str) -> None:
    df_top = df.head(15)

    print("")
    print("=================================")
    print(f"Top {len(df_top)} - {descriptor}")
    LOGGER.info(f"Top {len(df_top)} - {descriptor}")
    print("=================================")
    print(df_top)
    LOGGER.info(df_top)
    print("=================================")
    print("")


def save_to_excel(df: list[pd.DataFrame], file_name: str) -> None:
    base_path = r"c:\temp"

    if os.path.exists(r"\\busse\purchasing\Physical Inventory\2022\discrepancies"):
        base_path = r"\\busse\purchasing\Physical Inventory\2022\discrepancies"
        print("Using network path for output file.")
        print(f"Path: {base_path}")

    save_path = os.path.join(
        base_path, f"{file_name}_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
    )

    with pd.ExcelWriter(save_path, engine="xlsxwriter") as writer:
        sheet_names = {
            0: "All",
            1: "Sorted By Largest",
            2: "Sorted By Smallest",
        }
        for i, df in enumerate(df):
            try:
                df.to_excel(writer, sheet_name=sheet_names[i], index=False)
            except KeyError:
                df.to_excel(writer, sheet_name=f"Sheet{i+1}", index=False)


if __name__ == "__main__":
    rows, non_zero_rows = main()

    df_all = pd.DataFrame(rows)
    df_non_zero = pd.DataFrame(non_zero_rows)

    df_non_zero_largest = sort_largest_value(df_non_zero)
    df_non_zero_smallest = sort_smallest_value(df_non_zero)

    LOGGER.info("start")
    print_metrics(df_all)
    print_top(df_non_zero_largest, "largest")
    print_top(df_non_zero_smallest, "smallest")
    LOGGER.info("end")

    save_to_excel([df_all, df_non_zero_largest, df_non_zero_smallest], FILE_NAME)
