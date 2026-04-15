"""
Config Loader — reads Config_Data.xlsx 'Details' sheet.
Returns a dict with all configuration keys.

ADDED: 'inputPath' key — path to the Resource Effort input template
       (replaces hardcoded templatePath logic in projection_engine).
"""

import openpyxl
from datetime import datetime


class ConfigLoader:

    STRING_KEYS = {
        "kmgUrl", "gmailUrl", "timetrackEmail", "openMail",
        "outlookEmail", "password", "ToMail", "CcMail",
        "SubjectMail", "TextBoxMail", "downloadFilePath",
        "dateFilter", "templatePath", "inputPath", "estimationTab",
        "existingFilePath", "exportMode", "holidayPdf",
    }
    ARRAY_KEYS   = {"users", "projects", "months", "client", "billingStatus"}
    NUMERIC_KEYS = {"dailyLogFilter", "hrsPerDay"}

    @staticmethod
    def load(config_path: str) -> dict:
        wb = openpyxl.load_workbook(config_path, data_only=True)
        ws = wb["Details"]
        config = {}

        for row in ws.iter_rows(min_row=2, max_col=2):
            key_cell = row[0].value
            val_cell = row[1].value
            if key_cell is None:
                continue
            key = str(key_cell).strip()

            if isinstance(val_cell, datetime):
                config[key] = val_cell
                continue

            value = str(val_cell).strip() if val_cell is not None else ""
            if hasattr(val_cell, "text"):
                value = val_cell.text

            if key in ConfigLoader.ARRAY_KEYS:
                config[key] = [] if (not value or value.lower() == "select all") else [v.strip() for v in value.split(",")]
            elif key in ConfigLoader.NUMERIC_KEYS:
                try:
                    config[key] = float(value)
                except (ValueError, TypeError):
                    config[key] = 0
            else:
                config[key] = value

        wb.close()

        config.setdefault("hrsPerDay", 8)
        config.setdefault("exportMode", "auto")
        config.setdefault("estimationTab", "")
        config.setdefault("templatePath", "")
        config.setdefault("inputPath", "")
        config.setdefault("existingFilePath", "")
        config.setdefault("holidayPdf", "Holiday List 2026.pdf")
        return config
