import pandas as pd
import re
from datetime import datetime

def process_calendar_data(calendar_file, lookup_file):
    """
    Process calendar and lookup data from file-like objects.
    Args:
        calendar_file: file-like object for calendar CSV/Excel
        lookup_file: file-like object for organisation lookup Excel
    Returns:
        pd.DataFrame: Processed calendar DataFrame with organisation type and region
    """
    import io

    # Try reading as CSV, fallback to Excel
    try:
        calendar_df = pd.read_csv(calendar_file)
    except Exception:
        try:
            calendar_file.seek(0)
        except Exception:
            pass
        calendar_df = pd.read_excel(calendar_file)

    # Ensure required columns exist, fill missing with None
    required_columns = ["Date", "Required Attendees"]
    for col in required_columns:
        if col not in calendar_df.columns:
            calendar_df[col] = None

    # Convert 'Date' column to datetime
    calendar_df["Date"] = pd.to_datetime(calendar_df["Date"], dayfirst=True, errors="coerce")

    # Add a 'Quarter' column based on the 'Date'
    calendar_df["Quarter"] = calendar_df["Date"].dt.to_period("Q")

    # Convert date columns to datetime

    # (Removed Start Date/End Date logic; now only using 'Date')

    # Load organisation lookup from Excel
    lookup_df = pd.read_excel(lookup_file, sheet_name=None, engine="openpyxl")

    # Combine all sheets into one lookup table
    lookup_combined = pd.DataFrame()
    for sheet_name, sheet_df in lookup_df.items():
        sheet_df.columns = sheet_df.columns.str.strip()
        if "Type of organisation" in sheet_df.columns or "Type of Organisation" in sheet_df.columns:
            sheet_df["Type of Organisation"] = sheet_df.get("Type of organisation", sheet_df.get("Type of Organisation"))
            sheet_df["Region"] = sheet_df.apply(
                lambda row: ", ".join([
                    region for region in [
                        "Northern | Te Tai Tokerau",
                        "Midland | Te Manawa Taki",
                        "Central | Te Ikaroa",
                        "South Island | Te Waipounamu"
                    ] if str(row.get(region)).strip().lower() == "x"
                ]),
                axis=1
            )
            sheet_df = sheet_df[["Type of Organisation", "Region"]]
            lookup_combined = pd.concat([lookup_combined, sheet_df], ignore_index=True)

    # Extract domains from attendees
    def extract_domains(attendee_str):
        if pd.isna(attendee_str):
            return []
        emails = re.findall(r'[\w\.-]+@[\w\.-]+', attendee_str)
        domains = [email.split('@')[-1].lower() for email in emails]
        return domains

    calendar_df["Domains"] = calendar_df["Required Attendees"].apply(extract_domains)

    # Match organisation type and region from lookup using domain
    def match_lookup(domains):
        for domain in domains:
            match = lookup_combined[lookup_combined.apply(lambda row: domain in str(row.to_string()), axis=1)]
            if not match.empty:
                return match.iloc[0]["Type of Organisation"], match.iloc[0]["Region"]
        return None, None

    calendar_df[["Type of Organisation", "Region"]] = calendar_df["Domains"].apply(
        lambda domains: pd.Series(match_lookup(domains))
    )

    return calendar_df