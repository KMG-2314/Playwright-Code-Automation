"""
Calendar Utilities — week boundaries, holiday loading, working-day calculations.
Uses Sun–Sat week definition (matching existing ResourceEffortUpdater).
"""

import calendar
from datetime import date, timedelta
from typing import List, Dict, Set, Tuple

import pandas as pd
import pdfplumber


class CalendarUtils:
    """
    Provides calendar operations for month-based weekly resource planning.

    Week definition: Sun–Sat blocks within a single month.
      W1 starts on the 1st of the month (even if mid-week).
      W5 may be partial (ends on last day of month).
    """

    @staticmethod
    def get_week_boundaries(year: int, month: int) -> List[Tuple[int, date, date]]:
        """
        Returns list of (week_num, start_date, end_date) for the month.

        Week boundaries follow Sun–Sat pattern, clamped to month boundaries.
        Example for April 2026 (Apr 1 = Wednesday):
          W1: Apr 1 (Wed) – Apr 4 (Sat)
          W2: Apr 5 (Sun) – Apr 11 (Sat)
          W3: Apr 12 (Sun) – Apr 18 (Sat)
          W4: Apr 19 (Sun) – Apr 25 (Sat)
          W5: Apr 26 (Sun) – Apr 30 (Thu)
        """
        first_day = date(year, month, 1)
        last_day = date(year, month, calendar.monthrange(year, month)[1])

        # Find the first Saturday on or after the 1st
        # weekday(): Mon=0 ... Sun=6
        first_day_wd = first_day.weekday()  # 0=Mon
        # Sun=6 in Python's weekday(), Sat=5
        # Days until next Saturday (inclusive of first_day if it's a Saturday)
        days_to_sat = (5 - first_day_wd) % 7
        first_sat = first_day + timedelta(days=days_to_sat)

        weeks = []

        # W1: month_start to first Saturday (or end of month)
        w1_end = min(first_sat, last_day)
        weeks.append((1, first_day, w1_end))

        # Subsequent weeks: Sunday to Saturday
        week_num = 2
        current_sun = first_sat + timedelta(days=1)

        while current_sun <= last_day:
            current_sat = current_sun + timedelta(days=6)
            w_end = min(current_sat, last_day)
            weeks.append((week_num, current_sun, w_end))
            week_num += 1
            current_sun = current_sat + timedelta(days=1)

        return weeks

    @staticmethod
    def date_to_week_number(target_date: date, month: int, year: int) -> int:
        """
        Maps a specific date to W1–W5 for the given month.
        Returns the week number (1-based).
        """
        if isinstance(target_date, pd.Timestamp):
            target_date = target_date.date()

        weeks = CalendarUtils.get_week_boundaries(year, month)
        for week_num, w_start, w_end in weeks:
            if w_start <= target_date <= w_end:
                return week_num

        # If date is before the month, return W1; if after, return last week
        if target_date < date(year, month, 1):
            return 1
        return weeks[-1][0]

    @staticmethod
    def get_working_days_per_week(
        year: int, month: int, holiday_map: Dict[str, Set[date]], location: str = "Gurgaon"
    ) -> Dict[int, int]:
        """
        Returns {week_num: working_days} for each week of the month.
        holiday_map is { "Gurgaon": set(dates), "Kolkata": set(dates) }
        """
        # Specific holidays for this location
        loc_holidays = set()
        if location.lower() == "kolkata":
            loc_holidays = holiday_map.get("Kolkata", set())
        else:
            loc_holidays = holiday_map.get("Gurgaon", set())

        weeks = CalendarUtils.get_week_boundaries(year, month)
        result = {}

        for week_num, w_start, w_end in weeks:
            working = 0
            d = w_start
            while d <= w_end:
                # Mon(0) to Fri(4) = weekday
                if d.weekday() < 5 and d not in loc_holidays:
                    working += 1
                d += timedelta(days=1)
            result[week_num] = working

        return result

    @staticmethod
    def get_total_weeks(year: int, month: int) -> int:
        """Returns the number of weeks (4 or 5) in the given month."""
        weeks = CalendarUtils.get_week_boundaries(year, month)
        return len(weeks)

    @staticmethod
    def enddate_to_actual_weeks(
        end_date: date, month: int, year: int
    ) -> Tuple[List[int], List[int]]:
        """
        Given endDate from config, returns (actual_weeks, projected_weeks).

        Example: endDate = Apr 10 (in W2)
          actual_weeks = [1, 2]
          projected_weeks = [3, 4, 5]
        """
        if isinstance(end_date, pd.Timestamp):
            end_date = end_date.date()

        end_week = CalendarUtils.date_to_week_number(end_date, month, year)
        total = CalendarUtils.get_total_weeks(year, month)

        actual = list(range(1, end_week + 1))
        projected = list(range(end_week + 1, total + 1))

        return actual, projected

    @staticmethod
    def load_holidays_from_pdf(pdf_path: str) -> Dict[str, Set[date]]:
        """
        Parses the holiday PDF and returns separate holiday sets for Gurgaon and Kolkata.
        Expects a table where column 4 is Gurgaon and column 5 is Kolkata.
        Checkmark '√' (\u221a) indicates a holiday.
        """
        holidays = {"Gurgaon": set(), "Kolkata": set()}
        try:
            import pdfplumber
            with pdfplumber.open(pdf_path) as pdf:
                for page in pdf.pages:
                    table = page.extract_table()
                    if not table: continue
                    for row in table:
                        if not row or len(row) < 6: continue
                        date_str = (row[1] or "").strip()
                        if not date_str or date_str.lower() in ["date", "s. no."]: continue

                        try:
                            # Try to parse date
                            dt = pd.to_datetime(date_str, dayfirst=True).date()

                            # Check Gurgaon (index 4)
                            if "\u221a" in str(row[4] or ""):
                                holidays["Gurgaon"].add(dt)
                            # Check Kolkata (index 5)
                            if "\u221a" in str(row[5] or ""):
                                holidays["Kolkata"].add(dt)
                        except:
                            continue
            print(f"  [HOLIDAYS] Loaded {len(holidays['Gurgaon'])} for Gurgaon, {len(holidays['Kolkata'])} for Kolkata")
        except Exception as e:
            print(f"  [ERROR] PDF Parsing failed: {e}")
        return holidays

    @staticmethod
    def week_col_letter(week_num: int) -> str:
        """Maps week number (1-5) to Excel column letter (G-K)."""
        return {1: "G", 2: "H", 3: "I", 4: "J", 5: "K"}.get(week_num, "G")

    @staticmethod
    def week_col_index(week_num: int) -> int:
        """Maps week number (1-5) to Excel column index (7-11)."""
        return week_num + 6  # W1=col7(G), W2=col8(H), ...
