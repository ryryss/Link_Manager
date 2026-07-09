"""
data_manager.py
---------------

Responsible for all data operations:
- Loading data from Excel
- Saving data to Excel
- Importing CSV / Excel files
- Adding rows
- Updating cells
- Deleting rows
- Searching data

This module does NOT contain any GUI code.

Data structure:

    self.columns = [
        "Title",
        "Link",
        "Is Claimable",
        "Valid until"
    ]

    self.rows = [
        {
            "Title": "...",
            "Link": "...",
            "Is Claimable": "Yes",
            "Valid until": "..."
        }
    ]

Storage:
    Default file: link.xlsx
"""

import os
import sys
from typing import List, Dict, Tuple

import pandas as pd


import os

def get_app_path():
    """
    Get the directory where the exe is located.
    Works for both:
    - Running from Python source
    - Running as PyInstaller exe
    """

    if getattr(sys, "frozen", False):
        # Running as compiled exe
        return os.path.dirname(sys.executable)

    else:
        # Running from Python
        return os.path.dirname(
            os.path.abspath(__file__)
        )

class DataManager:

    DEFAULT_STORAGE_FILE = "links.xlsx"


    def __init__(self, storage_path=None):

        if storage_path:
            self.storage_path = storage_path

        else:
            self.storage_path = os.path.join(
                get_app_path(),
                self.DEFAULT_STORAGE_FILE
            )


        self.columns = []
        self.rows = []

        self.load()



    # -------------------------------------------------
    # Persistence
    # -------------------------------------------------

    def load(self):
        """
        Load existing data from Excel file.

        If the file does not exist,
        initialize an empty database.
        """

        if os.path.exists(self.storage_path):

            df = pd.read_excel(
                self.storage_path,
                dtype=str
            ).fillna("")


            self.columns = list(
                df.columns
            )


            self.rows = df.to_dict(
                orient="records"
            )


        else:

            self.columns = []
            self.rows = []



    def save(self):
        """Save data to Excel."""

        columns = list(self.columns)

        if "Status" not in columns:
            columns.append("Status")

        df = pd.DataFrame(self.rows)

        df = df.reindex(columns=columns)

        df.to_excel(
            self.storage_path,
            index=False
        )


    # -------------------------------------------------
    # File Import
    # -------------------------------------------------

    @staticmethod
    def import_file(
        path: str
    ) -> Tuple[List[str], List[Dict[str, str]]]:
        """
        Import CSV or Excel file.

        Supported formats:
            .xlsx
            .xls
            .csv
            .tsv
            .txt

        Returns:
            columns:
                List of column names

            rows:
                List of dictionaries
        """

        ext = os.path.splitext(path)[1].lower()


        if ext in (
            ".xlsx",
            ".xls"
        ):

            df = pd.read_excel(
                path,
                dtype=str
            )


        elif ext in (
            ".csv",
            ".tsv",
            ".txt"
        ):

            separator = (
                "\t"
                if ext == ".tsv"
                else None
            )


            df = pd.read_csv(
                path,
                dtype=str,
                sep=separator,
                engine="python"
            )


        else:

            raise ValueError(
                f"Unsupported file type: {ext}"
            )



        df = df.fillna("")

        # Remove completely empty rows
        df = df[
            df.astype(str)
              .apply(
                  lambda row: row.str.strip().ne("").any(),
                  axis=1
              )
        ]
        columns = list(
            df.columns
        )


        rows = df.to_dict(
            orient="records"
        )


        return columns, rows

    # -------------------------------------------------
    # Create
    # -------------------------------------------------

    def add_rows(self, new_columns: List[str], new_rows: List[Dict[str, str]]):

        if not self.columns:
            self.columns = [c for c in new_columns if c != "Status"]

        existing_links = {
            str(row.get("Link", "")).strip()
            for row in self.rows
            if str(row.get("Link", "")).strip()
        }

        for row in new_rows:

            link = str(row.get("Link", "")).strip()

            if not link:
                continue

            if link in existing_links:
                continue

            aligned = {
                col: row.get(col, "")
                for col in self.columns
            }

            aligned["Status"] = ""

            self.rows.append(aligned)

            existing_links.add(link)

        self.save()


    # -------------------------------------------------
    # Update
    # -------------------------------------------------

    def update_cell(
        self,
        row_index: int,
        column: str,
        value: str
    ):
        """
        Update a single cell value.
        """

        if (
            0 <= row_index < len(self.rows)
            and column in self.columns
        ):

            self.rows[row_index][column] = value

            self.save()

    def update_status(self, row_index: int, status: str):

        if 0 <= row_index < len(self.rows):

            self.rows[row_index]["Status"] = status

            self.save()
    # -------------------------------------------------
    # Delete
    # -------------------------------------------------

    def delete_rows(
        self,
        row_indices: List[int]
    ):
        """
        Delete rows by index.
        """

        for idx in sorted(
            set(row_indices),
            reverse=True
        ):

            if (
                0 <= idx < len(self.rows)
            ):

                del self.rows[idx]


        self.save()



    # -------------------------------------------------
    # Search
    # -------------------------------------------------

    def search(
        self,
        keyword: str
    ) -> List[int]:
        """
        Search all columns.

        Matching:
            Case insensitive
            Partial match

        Returns:
            List of row indexes
        """

        if not keyword:

            return list(
                range(len(self.rows))
            )


        keyword = keyword.lower()


        result = []


        for index, row in enumerate(
            self.rows
        ):

            for value in row.values():

                if keyword in str(value).lower():

                    result.append(index)

                    break



        return result