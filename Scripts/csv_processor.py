"""
CSV Processor — loads Timesheet.csv, cleans data, returns structured actual hours.

FIXES:
  1. Normalize whitespace in owner names before LEAD_MAP lookup.
  2. Lead map now correctly merges Anushka/Ankush/Durva -> Teena Walia,
     Bhawesh -> Arun Kumar.
  3. NaN-safe processing.
"""

import re
import pandas as pd
from typing import Dict, Any
from Scripts.calendar_utils import CalendarUtils


class CSVProcessor:

    OWNER_CANONICAL_MAP = {
        "S Praveen Kumar": "Praveen Kumar",
    }

    LEAD_MAP = {
        "Anushka Gupta": "Teena Walia",
        "Ankush Ujjwal": "Teena Walia",
        "Durva Agarwal": "Teena Walia",
        "Bhawesh Pant": "Arun Kumar",
    }

    def load_and_process(self, csv_path: str, month: int, year: int) -> Dict[str, Any]:
        df = pd.read_csv(csv_path)

        # Clean Owner
        df["Owner"] = df["Owner"].astype(str).str.strip()
        df["Owner"] = df["Owner"].replace(self.OWNER_CANONICAL_MAP)
        # Normalise multiple spaces
        df["Owner"] = df["Owner"].apply(
            lambda x: re.sub(r"\s+", " ", x).strip() if pd.notna(x) and x != "nan" else ""
        )
        # Apply lead mapping
        df["Owner"] = df["Owner"].apply(lambda x: self.LEAD_MAP.get(x, x))
        df = df[df["Owner"].notna() & (df["Owner"] != "") & (df["Owner"] != "nan")]

        # Clean Daily Log
        df["Daily Log"] = pd.to_numeric(df["Daily Log"], errors="coerce").fillna(0)

        # Clean Timesheet Date
        df["Timesheet Date"] = (
            df["Timesheet Date"]
            .astype(str)
            .str.replace('="', "", regex=False)
            .str.replace('"', "", regex=False)
            .str.strip()
        )
        df["Timesheet Date"] = pd.to_datetime(
            df["Timesheet Date"], format="%m/%d/%Y", errors="coerce"
        )
        df = df.dropna(subset=["Timesheet Date"])

        df["Project"]    = df["Project"].apply(self.clean_allocation_name)
        df["Assignment"] = df["Assignment"].apply(self.clean_allocation_name)

        df["WeekNum"] = df["Timesheet Date"].apply(
            lambda d: CalendarUtils.date_to_week_number(d.date(), month, year)
        )

        agg = (
            df.groupby(["Owner", "Project", "Assignment", "WeekNum"], as_index=False)
            .agg({"Daily Log": "sum"})
        )

        result = {}
        for _, row in agg.iterrows():
            owner   = row["Owner"]
            project = row["Project"]
            task    = row["Assignment"]
            week    = int(row["WeekNum"])
            hours   = float(row["Daily Log"])

            if owner not in result:
                result[owner] = {"tasks": [], "total_actual": 0}

            task_entry = next(
                (t for t in result[owner]["tasks"]
                 if t["project"] == project and t["task"] == task),
                None,
            )
            if task_entry is None:
                task_entry = {"project": project, "task": task, "hours": {}}
                result[owner]["tasks"].append(task_entry)

            task_entry["hours"][week] = task_entry["hours"].get(week, 0) + hours

        for data in result.values():
            data["total_actual"] = sum(sum(t["hours"].values()) for t in data["tasks"])

        return result

    @staticmethod
    def clean_allocation_name(name: str) -> str:
        if not name or not isinstance(name, str) or name.lower() == "nan":
            return ""
        # 1. Remove "Quincy" prefix noise
        name = re.sub(r"(?i)^.*?quincy\s*[-_\s\u2013\u2014]?\s*", "", name)
        # 2. Remove leading numbers/symbols: "01_", "02-", "03 "
        name = re.sub(r"^[^\w]*\d+[\s\-_]*[\.\-_]*\s*", "", name)
        # 3. Replace separators with spaces (so digit tokens next to '_'/'-' become removable)
        name = re.sub(r"[-_()\[\]{}/\\&;,.]", " ", name)
        # 4. Remove standalone numbers in middle (after separator normalization)
        name = re.sub(r"\b\d+\b", "", name)
        # 5. Normalize whitespace and Title Case
        name = " ".join(name.split()).strip().title()
        return name

    @staticmethod
    def clean_project_name(name: str) -> str:
        return CSVProcessor.clean_allocation_name(name)

    @staticmethod
    def clean_task_name(name: str) -> str:
        return CSVProcessor.clean_allocation_name(name)
