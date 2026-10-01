"""
Skytecher Cold Email Automation - Excel Handler Module
======================================================
Handles loading, fuzzy column matching, email validation, row filtering,
real-time status updates back to the Excel file, and preview exports.
"""

import re
import datetime
from pathlib import Path
from typing import List, Dict, Any, Tuple, Optional
import pandas as pd
import openpyxl


# Standardized internal field names
FIELD_CLIENT_NAME = "client_name"
FIELD_COMPANY_NAME = "company_name"
FIELD_EMAIL = "email"
FIELD_INDUSTRY = "industry"
FIELD_WEBSITE = "website"
FIELD_CITY = "city"
FIELD_COUNTRY = "country"
FIELD_DESCRIPTION = "description"
FIELD_SUGGESTED_SERVICE = "suggested_service"
FIELD_HAS_WEBSITE = "has_website"
FIELD_LEAD_TYPE = "lead_type"
FIELD_LINKEDIN = "linkedin"
FIELD_INSTAGRAM = "instagram"
FIELD_MAPS_LINK = "maps_link"
FIELD_STATUS = "status"
FIELD_SENT_AT = "sent_date_time"
FIELD_ERROR_MESSAGE = "error_message"

# Comprehensive aliases for fuzzy column name resolution
# Supports BOTH the standard Skytecher format AND the user's custom lead-gen format:
#   Company, Category, Country, City, Email(s), Website, Has Website?,
#   Lead Type, Suggestion, LinkedIn, Instagram, Maps Link
COLUMN_ALIASES = {
    FIELD_CLIENT_NAME: [
        "client name", "client", "name", "contact name", "contact person",
        "full name", "lead name", "person", "recipient name", "first name"
    ],
    FIELD_COMPANY_NAME: [
        "company name", "company", "business name", "business", "organization",
        "org name", "firm", "client company", "account name"
    ],
    FIELD_EMAIL: [
        "email", "email(s)", "emails", "email id", "email address", "e-mail",
        "e-mail id", "contact email", "work email", "mail", "recipient email"
    ],
    FIELD_INDUSTRY: [
        "industry", "category", "sector", "business sector", "domain",
        "vertical", "niche", "field", "business category"
    ],
    FIELD_WEBSITE: [
        "website", "url", "web", "site", "web page", "domain name", "company url"
    ],
    FIELD_CITY: [
        "city", "location", "address", "state", "region", "town", "headquarters"
    ],
    FIELD_COUNTRY: [
        "country", "nation", "country name"
    ],
    FIELD_DESCRIPTION: [
        "company details / description", "company details", "description",
        "about", "company description", "about company", "details", "notes",
        "business description", "summary", "pain point", "pain points",
        "problem", "likely problem", "their likely problem or need",
        "problem or need", "challenges", "need"
    ],
    FIELD_SUGGESTED_SERVICE: [
        "suggested service", "suggested services", "suggestion", "service",
        "services", "recommended service", "pitch service", "target service",
        "offering"
    ],
    FIELD_HAS_WEBSITE: [
        "has website", "has website?", "website status", "has site"
    ],
    FIELD_LEAD_TYPE: [
        "lead type", "lead category", "lead source", "type", "prospect type"
    ],
    FIELD_LINKEDIN: [
        "linkedin", "linkedin url", "linkedin profile", "linkedin link"
    ],
    FIELD_INSTAGRAM: [
        "instagram", "instagram url", "instagram profile", "ig", "instagram link"
    ],
    FIELD_MAPS_LINK: [
        "maps link", "google maps", "maps url", "map link", "google maps link",
        "maps"
    ],
    FIELD_STATUS: [
        "status", "email status", "send status", "delivery status", "campaign status"
    ],
    FIELD_SENT_AT: [
        "sent date & time", "sent date and time", "sent at", "sent date",
        "date sent", "timestamp", "last sent"
    ],
    FIELD_ERROR_MESSAGE: [
        "error message", "error", "failure reason", "notes/errors", "log"
    ],
}

# RFC 5322 compliant practical regex for email validation
EMAIL_REGEX = re.compile(
    r"^[a-zA-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?(?:\.[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?)+$"
)


def normalize_header(header: str) -> str:
    """Normalize a column header string for robust fuzzy matching."""
    if not isinstance(header, str):
        return ""
    # Lowercase, replace slashes and underscores with spaces, collapse spaces
    cleaned = header.lower().strip()
    cleaned = re.sub(r"[/_\-]+", " ", cleaned)
    cleaned = re.sub(r"\s+", " ", cleaned)
    return cleaned.strip()


def is_valid_email(email: Any) -> bool:
    """Validate if an email is a non-empty string and conforms to RFC regex."""
    if not email or not isinstance(email, str):
        return False
    email = email.strip()
    if len(email) < 5 or "@" not in email:
        return False
    return bool(EMAIL_REGEX.match(email))


def extract_clean_email(raw_email: Any) -> str:
    """
    Extracts the first valid email address from a string that may contain
    multiple comma- or semicolon-separated emails, or angle brackets.
    Example: 'Contact <info@company.com>, sales@company.com' -> 'info@company.com'
    """
    if not raw_email or not isinstance(raw_email, str):
        return ""
    # Find all email pattern matches in the string
    matches = re.findall(
        r"[a-zA-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?(?:\.[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?)+",
        str(raw_email),
    )
    for m in matches:
        if is_valid_email(m):
            return m
    # Fallback to stripped string if already valid
    s = str(raw_email).strip().strip("<>()[]\"'")
    if is_valid_email(s):
        return s
    return ""


class ExcelHandler:
    """
    Manages loading, parsing, validating, and updating client Excel files.
    """

    def __init__(self, file_path: str | Path):
        self.file_path = Path(file_path).resolve()
        self.df: Optional[pd.DataFrame] = None
        self.column_mapping: Dict[str, str] = {}  # internal_name -> actual_col_name
        self.reverse_mapping: Dict[str, str] = {}  # actual_col_name -> internal_name

    def load_file(self) -> Tuple[bool, str]:
        """
        Loads the Excel file using openpyxl engine and builds column mappings.
        """
        if not self.file_path.exists():
            return False, f"File not found at: {self.file_path}"

        try:
            # Read all sheets, default to the first one
            self.df = pd.read_excel(self.file_path, engine="openpyxl")
            # Replace NaN/null with empty strings for text safety
            self.df = self.df.fillna("")
            self._detect_columns()
            return True, f"Successfully loaded {len(self.df)} rows from {self.file_path.name}"
        except Exception as e:
            return False, f"Failed to read Excel file: {str(e)}"

    def _detect_columns(self) -> None:
        """
        Auto-detects actual Excel column headers using the alias dictionary.
        """
        if self.df is None:
            return

        actual_cols = list(self.df.columns)
        self.column_mapping = {}
        self.reverse_mapping = {}

        for col in actual_cols:
            norm = normalize_header(str(col))
            matched = False
            for internal_key, aliases in COLUMN_ALIASES.items():
                if internal_key in self.column_mapping:
                    continue  # already mapped this field
                for alias in aliases:
                    if norm == alias or alias in norm:
                        self.column_mapping[internal_key] = col
                        self.reverse_mapping[col] = internal_key
                        matched = True
                        break
                if matched:
                    break

        # Ensure required tracking columns exist in the DataFrame
        for tracking_field, default_col_name in [
            (FIELD_STATUS, "Status"),
            (FIELD_SENT_AT, "Sent Date & Time"),
            (FIELD_ERROR_MESSAGE, "Error Message"),
        ]:
            if tracking_field not in self.column_mapping:
                self.df[default_col_name] = ""
                self.column_mapping[tracking_field] = default_col_name
                self.reverse_mapping[default_col_name] = tracking_field

    def get_column_mapping_summary(self) -> Dict[str, str]:
        """Returns readable report of how Excel columns were mapped."""
        return {
            internal: self.column_mapping.get(internal, "Not Found")
            for internal in [
                FIELD_CLIENT_NAME,
                FIELD_COMPANY_NAME,
                FIELD_EMAIL,
                FIELD_INDUSTRY,
                FIELD_WEBSITE,
                FIELD_CITY,
                FIELD_COUNTRY,
                FIELD_DESCRIPTION,
                FIELD_SUGGESTED_SERVICE,
                FIELD_HAS_WEBSITE,
                FIELD_LEAD_TYPE,
                FIELD_LINKEDIN,
                FIELD_INSTAGRAM,
                FIELD_MAPS_LINK,
                FIELD_STATUS,
            ]
        }

    def get_clients(self) -> List[Dict[str, Any]]:
        """
        Extracts all client records from the dataframe into structured dicts.
        Preserves original 0-indexed row index and 1-indexed Excel row number.
        Automatically combines Country + City into the city field for location context.
        Uses Company Name as fallback for Client Name if missing.
        """
        if self.df is None:
            return []

        clients = []
        for idx, row in self.df.iterrows():
            # Helper to safely extract a mapped field
            def _get(field: str) -> str:
                col = self.column_mapping.get(field, "")
                return str(row.get(col, "")).strip() if col else ""

            company_name = _get(FIELD_COMPANY_NAME)
            client_name = _get(FIELD_CLIENT_NAME)
            city = _get(FIELD_CITY)
            country = _get(FIELD_COUNTRY)
            raw_email = _get(FIELD_EMAIL)
            clean_email = extract_clean_email(raw_email) or raw_email

            # If client name is blank, use company name as the greeting name
            if not client_name:
                client_name = company_name

            # Combine City + Country for richer location context
            location = city
            if city and country:
                location = f"{city}, {country}"
            elif country and not city:
                location = country

            client_dict = {
                "_row_index": idx,
                # In Excel, row 1 is header, so row 0 in df corresponds to Excel row 2
                "_excel_row_num": idx + 2,
                FIELD_CLIENT_NAME: client_name,
                FIELD_COMPANY_NAME: company_name,
                FIELD_EMAIL: clean_email,
                "_raw_email": raw_email,
                FIELD_INDUSTRY: _get(FIELD_INDUSTRY),
                FIELD_WEBSITE: _get(FIELD_WEBSITE),
                FIELD_CITY: location,
                FIELD_COUNTRY: country,
                FIELD_DESCRIPTION: _get(FIELD_DESCRIPTION),
                FIELD_SUGGESTED_SERVICE: _get(FIELD_SUGGESTED_SERVICE),
                FIELD_HAS_WEBSITE: _get(FIELD_HAS_WEBSITE),
                FIELD_LEAD_TYPE: _get(FIELD_LEAD_TYPE),
                FIELD_LINKEDIN: _get(FIELD_LINKEDIN),
                FIELD_INSTAGRAM: _get(FIELD_INSTAGRAM),
                FIELD_MAPS_LINK: _get(FIELD_MAPS_LINK),
                FIELD_STATUS: _get(FIELD_STATUS),
                FIELD_SENT_AT: _get(FIELD_SENT_AT),
                FIELD_ERROR_MESSAGE: _get(FIELD_ERROR_MESSAGE),
            }
            clients.append(client_dict)
        return clients

    def update_row_status(
        self,
        row_index: int,
        status: str,
        sent_at: Optional[str] = None,
        error_msg: str = "",
    ) -> bool:
        """
        Updates the row status in both memory (DataFrame) and immediately
        persists it to the Excel file using openpyxl for live safety.
        """
        if self.df is None or row_index < 0 or row_index >= len(self.df):
            return False

        if sent_at is None and status == "Sent":
            sent_at = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        elif sent_at is None:
            sent_at = ""

        # Update in memory dataframe
        status_col = self.column_mapping.get(FIELD_STATUS, "Status")
        sent_col = self.column_mapping.get(FIELD_SENT_AT, "Sent Date & Time")
        err_col = self.column_mapping.get(FIELD_ERROR_MESSAGE, "Error Message")

        self.df.at[row_index, status_col] = status
        self.df.at[row_index, sent_col] = sent_at
        self.df.at[row_index, err_col] = error_msg

        # Write directly to Excel file using openpyxl for immediate persistence
        try:
            wb = openpyxl.load_workbook(self.file_path)
            ws = wb.active

            # Find column indices (1-indexed) in sheet
            header_map = {}
            for col_idx in range(1, ws.max_column + 1):
                cell_val = ws.cell(row=1, column=col_idx).value
                if cell_val is not None:
                    header_map[str(cell_val).strip()] = col_idx

            # Add missing tracking headers if needed
            for field, col_name in [
                (FIELD_STATUS, status_col),
                (FIELD_SENT_AT, sent_col),
                (FIELD_ERROR_MESSAGE, err_col),
            ]:
                if col_name not in header_map:
                    new_col_idx = ws.max_column + 1
                    ws.cell(row=1, column=new_col_idx, value=col_name)
                    header_map[col_name] = new_col_idx

            excel_row = row_index + 2  # header is row 1
            ws.cell(row=excel_row, column=header_map[status_col], value=status)
            ws.cell(row=excel_row, column=header_map[sent_col], value=sent_at)
            ws.cell(row=excel_row, column=header_map[err_col], value=error_msg)

            wb.save(self.file_path)
            wb.close()
            return True
        except Exception as e:
            print(f"[!] Warning: Could not update Excel file on disk: {e}")
            return False

    def export_preview(
        self, preview_records: List[Dict[str, Any]], output_path: Path
    ) -> bool:
        """
        Exports generated email previews to an Excel spreadsheet for inspection.
        """
        try:
            preview_df = pd.DataFrame(preview_records)
            preview_df.to_excel(output_path, index=False, engine="openpyxl")
            return True
        except Exception as e:
            print(f"[!] Error exporting preview to Excel: {e}")
            return False

    @staticmethod
    def generate_sample_excel(output_path: Path) -> Path:
        """
        Generates a high-quality sample Excel file with 10 realistic clients
        using the user's custom lead-gen column format:
        Company, Category, Country, City, Email(s), Website, Has Website?,
        Lead Type, Suggestion, LinkedIn, Instagram, Maps Link
        """
        sample_data = [
            {
                "Company": "Apex Dental Clinic",
                "Category": "Healthcare",
                "Country": "United States",
                "City": "Austin",
                "Email(s)": "sjenkins@apexdentalcare.org",
                "Website": "https://apexdentalcare.org",
                "Has Website?": "Yes",
                "Lead Type": "Warm Lead",
                "Suggestion": "Website Design & Development",
                "LinkedIn": "https://linkedin.com/company/apex-dental",
                "Instagram": "https://instagram.com/apexdental",
                "Maps Link": "https://maps.google.com/?q=Apex+Dental+Austin",
            },
            {
                "Company": "UrbanVibe Apparel",
                "Category": "E-Commerce / Retail",
                "Country": "United States",
                "City": "New York",
                "Email(s)": "marcus@urbanvibeapparel.com",
                "Website": "https://urbanvibeapparel.com",
                "Has Website?": "Yes",
                "Lead Type": "Cold Lead",
                "Suggestion": "",
                "LinkedIn": "https://linkedin.com/company/urbanvibe",
                "Instagram": "https://instagram.com/urbanvibeapparel",
                "Maps Link": "",
            },
            {
                "Company": "PaySwift Financial",
                "Category": "FinTech / SaaS",
                "Country": "United States",
                "City": "San Francisco",
                "Email(s)": "elena.rostova@payswift.io",
                "Website": "https://payswift.io",
                "Has Website?": "Yes",
                "Lead Type": "Warm Lead",
                "Suggestion": "Mobile App Development",
                "LinkedIn": "https://linkedin.com/company/payswift",
                "Instagram": "",
                "Maps Link": "",
            },
            {
                "Company": "Horizon Logistics Hub",
                "Category": "Logistics & Supply Chain",
                "Country": "United States",
                "City": "Chicago",
                "Email(s)": "david.chen@horizonlogix.net",
                "Website": "https://horizonlogix.net",
                "Has Website?": "Yes",
                "Lead Type": "Cold Lead",
                "Suggestion": "",
                "LinkedIn": "",
                "Instagram": "",
                "Maps Link": "https://maps.google.com/?q=Horizon+Logistics+Chicago",
            },
            {
                "Company": "Aura Luxe Interior Design",
                "Category": "Luxury Interior Design",
                "Country": "United States",
                "City": "Miami",
                "Email(s)": "chloe@auraluxedesign.com",
                "Website": "https://auraluxedesign.com",
                "Has Website?": "Yes",
                "Lead Type": "Warm Lead",
                "Suggestion": "Branding & UI/UX Design",
                "LinkedIn": "https://linkedin.com/company/aura-luxe",
                "Instagram": "https://instagram.com/auraluxedesign",
                "Maps Link": "",
            },
            {
                "Company": "Sterling Capital Partners",
                "Category": "Finance & Wealth Management",
                "Country": "United States",
                "City": "Boston",
                "Email(s)": "rsterling@sterlingcapllc.com",
                "Website": "https://sterlingcapllc.com",
                "Has Website?": "Yes",
                "Lead Type": "Cold Lead",
                "Suggestion": "IT Consulting & Support",
                "LinkedIn": "https://linkedin.com/company/sterling-capital",
                "Instagram": "",
                "Maps Link": "",
            },
            {
                "Company": "FreshBite Cloud Kitchens",
                "Category": "Food & Beverage / Tech",
                "Country": "United States",
                "City": "Seattle",
                "Email(s)": "amara@freshbitekitchens.co",
                "Website": "https://freshbitekitchens.co",
                "Has Website?": "Yes",
                "Lead Type": "Warm Lead",
                "Suggestion": "",
                "LinkedIn": "",
                "Instagram": "https://instagram.com/freshbitekitchens",
                "Maps Link": "",
            },
            {
                "Company": "Summit Peak Construction",
                "Category": "Commercial Construction",
                "Country": "United States",
                "City": "Denver",
                "Email(s)": "mgallagher@summitpeakbuilders.com",
                "Website": "",
                "Has Website?": "No",
                "Lead Type": "Cold Lead",
                "Suggestion": "",
                "LinkedIn": "",
                "Instagram": "",
                "Maps Link": "https://maps.google.com/?q=Summit+Peak+Denver",
            },
            {
                "Company": "O'Connor Legal Group",
                "Category": "Legal Services",
                "Country": "Ireland",
                "City": "Dublin",
                "Email(s)": "liam@oconnorlaw.ie",
                "Website": "https://oconnorlaw.ie",
                "Has Website?": "Yes",
                "Lead Type": "Cold Lead",
                "Suggestion": "Digital Marketing & SEO",
                "LinkedIn": "https://linkedin.com/company/oconnor-legal",
                "Instagram": "",
                "Maps Link": "",
            },
            {
                "Company": "NovaFit Studios",
                "Category": "Fitness & Wellness",
                "Country": "United States",
                "City": "Los Angeles",
                "Email(s)": "sophia@novafitstudios.com",
                "Website": "https://novafitstudios.com",
                "Has Website?": "Yes",
                "Lead Type": "Warm Lead",
                "Suggestion": "Software / Custom Development",
                "LinkedIn": "https://linkedin.com/company/novafit-studios",
                "Instagram": "https://instagram.com/novafitstudios",
                "Maps Link": "",
            },
        ]

        df = pd.DataFrame(sample_data)
        # Add empty tracking columns
        df["Status"] = ""
        df["Sent Date & Time"] = ""
        df["Error Message"] = ""

        df.to_excel(output_path, index=False, engine="openpyxl")
        return output_path

    @staticmethod
    def generate_template_excel(output_path: Path) -> Path:
        """
        Generates a blank Excel template with the exact column headers
        the user should fill in. Contains 1 example row for reference.
        """
        template_data = [
            {
                "Company": "Example Corp",
                "Category": "Technology",
                "Country": "United States",
                "City": "New York",
                "Email(s)": "hello@examplecorp.com",
                "Website": "https://examplecorp.com",
                "Has Website?": "Yes",
                "Lead Type": "Cold Lead",
                "Suggestion": "Website Design & Development",
                "LinkedIn": "https://linkedin.com/company/examplecorp",
                "Instagram": "https://instagram.com/examplecorp",
                "Maps Link": "https://maps.google.com/?q=Example+Corp",
            },
        ]

        df = pd.DataFrame(template_data)
        df["Status"] = ""
        df["Sent Date & Time"] = ""
        df["Error Message"] = ""

        df.to_excel(output_path, index=False, engine="openpyxl")

        # Style the header row for better visibility
        try:
            from openpyxl.styles import Font, PatternFill, Alignment

            wb = openpyxl.load_workbook(output_path)
            ws = wb.active
            header_fill = PatternFill(start_color="1E3A5F", end_color="1E3A5F", fill_type="solid")
            header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
            for col in range(1, ws.max_column + 1):
                cell = ws.cell(row=1, column=col)
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")
                # Auto-fit column width (rough estimate)
                ws.column_dimensions[cell.column_letter].width = max(len(str(cell.value)) + 4, 15)
            wb.save(output_path)
            wb.close()
        except Exception:
            pass  # Styling is optional; file still works without it

        return output_path
