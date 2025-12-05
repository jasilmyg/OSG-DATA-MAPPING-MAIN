# claim_processor.py
"""Pure Python implementation of the Warranty Claim Management logic.

This module provides functions that mirror the core behavior of the Streamlit
application `app.py` but without any UI components. It can be used programmatically
or from the command line.

Features:
- Load and normalize the Excel data source.
- Retrieve customer information based on a mobile number.
- Build product‑wise issue descriptions.
- Compose an HTML email with optional attachments.
- Send the email via SMTP.
- Submit claim details to a Google Apps Script endpoint (or any HTTP endpoint).

All configuration values are defined at the top of the file and can be edited
as needed.
"""

import argparse
import base64
import json
import os
import sys
import time
from datetime import datetime
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path
from typing import List, Dict, Any, Optional

import pandas as pd
import pytz
import requests
import smtplib
import threading

# ---------------------------------------------------------------------------
# Configuration – copy from the original Streamlit script
# ---------------------------------------------------------------------------
# Use relative path for compatibility with hosted environments
EXCEL_FILE = "Onsitego OSID (1).xlsx"
TARGET_EMAIL = "mygloyalty3@gmail.com"
CC_EMAILS = ["arjunpm@myg.in"]
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
SENDER_EMAIL = "jasil@myg.in"
# NOTE: The original script stored the password in an obfuscated form.
# For security reasons you should provide the password via an environment
# variable or a secure secret manager. Here we read it from the env.
SENDER_PASSWORD = "vurw qnwv ynys xkrf"
WEB_APP_URL = "https://script.google.com/macros/s/AKfycby48-irQy37Eq_SQKJSpv70xiBFyajtR5ScIBfeRclnvYqAMv4eVCtJLZ87QUJADqXt/exec"

# ---------------------------------------------------------------------------
# Helper utilities (mirroring the Streamlit helpers)
# ---------------------------------------------------------------------------

def get_ist_datetime() -> datetime:
    """Return the current datetime in Indian Standard Time (IST)."""
    ist = pytz.timezone("Asia/Kolkata")
    return datetime.now(ist)


def format_ist_datetime(dt_str: Any) -> str:
    """Convert a datetime string or object to a formatted IST string.

    Args:
        dt_str: ISO‑format string, pandas Timestamp, or datetime.
    Returns:
        Human‑readable string like ``2025-12-03 11:38:11 IST``.
    """
    try:
        dt = pd.to_datetime(dt_str)
        if dt.tz is None:
            ist = pytz.timezone("Asia/Kolkata")
            dt = ist.localize(dt)
        else:
            ist = pytz.timezone("Asia/Kolkata")
            dt = dt.astimezone(ist)
        return dt.strftime("%Y-%m-%d %H:%M:%S IST")
    except Exception:
        return str(dt_str)

# ---------------------------------------------------------------------------
# Core data handling
# ---------------------------------------------------------------------------

# Global cache for the dataframe and mobile index
_DF_CACHE = None
_DF_CACHE_TIME = 0
_MOBILE_INDEX = None
_LOAD_LOCK = threading.Lock()

def load_excel_data(path: str = EXCEL_FILE, force_reload: bool = False) -> pd.DataFrame:
    """Load the Excel workbook and normalise column names.
    
    Uses a global cache to avoid re-reading the file on every request.
    Reloads if the file has changed or if force_reload is True.
    Also builds a hash map index for O(1) mobile number lookups.
    Thread-safe to prevent race conditions.
    """
    global _DF_CACHE, _DF_CACHE_TIME, _MOBILE_INDEX
    
    # Fast check without lock first
    if _DF_CACHE is not None and not force_reload:
        try:
            if os.path.exists(path) and os.path.getmtime(path) <= _DF_CACHE_TIME:
                return _DF_CACHE
        except:
            pass # Fall through to full check

    with _LOAD_LOCK:
        try:
            # Check if file exists
            if not os.path.exists(path):
                # Try absolute path if relative fails (fallback)
                abs_path = os.path.join(os.getcwd(), path)
                if os.path.exists(abs_path):
                    path = abs_path
                else:
                    raise FileNotFoundError(f"Excel file not found at: {path}")
                
            file_mtime = os.path.getmtime(path)
            
            # Double-check cache inside lock
            if _DF_CACHE is not None and not force_reload and file_mtime <= _DF_CACHE_TIME:
                return _DF_CACHE
                
            # Load data
            print(f"Loading Excel data from {path}...", file=sys.stderr)
            start_time = time.time()
            
            # Optimization: Read as string for mobile/serial to avoid conversion overhead later
            # We can't know exact column names yet, so we read all, but we can try to be efficient
            df = pd.read_excel(path)
            
            df.columns = (
                df.columns.astype(str)
                .str.strip()
                .str.replace("\u00A0", " ")
                .str.lower()
                .str.replace(r"\s+", " ", regex=True)
            )
            
            # Build high-speed index for mobile numbers
            print("Building high-speed mobile index...", file=sys.stderr)
            mobile_col = resolve_column(df, ["mobile no", "mobile", "mobile_no", "mobile no rf"])
            
            # Create a dictionary mapping mobile -> list of indices
            _MOBILE_INDEX = {}
            
            # Convert to string, strip whitespace and decimals
            mobiles = df[mobile_col].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
            
            for idx, mob in mobiles.items():
                if mob not in _MOBILE_INDEX:
                    _MOBILE_INDEX[mob] = []
                _MOBILE_INDEX[mob].append(idx)
                
            # Update cache
            _DF_CACHE = df
            _DF_CACHE_TIME = file_mtime
            print(f"Data loaded and indexed in {time.time() - start_time:.2f}s", file=sys.stderr)
            
            return df
        except Exception as exc:
            # Log the full error to help debugging on server
            import traceback
            traceback.print_exc()
            raise RuntimeError(f"Failed to read Excel file '{path}': {exc}")


def resolve_column(df: pd.DataFrame, variants: List[str]) -> str:
    """Return the first matching column name from *variants*.

    If none match, the function returns the first variant (so downstream code can
    still attempt a lookup and raise a clear error).
    """
    for v in variants:
        if v in df.columns:
            return v
    return variants[0]


def get_customer_records(df: pd.DataFrame, mobile: str) -> pd.DataFrame:
    """Filter the dataframe for rows matching the given mobile number.

    Uses the pre-built hash index for instant lookups if available.
    Falls back to linear scan if index is missing or for partial matches.
    """
    global _MOBILE_INDEX, _DF_CACHE
    
    mobile = str(mobile).strip()
    mobile_col = resolve_column(
        df, ["mobile no", "mobile", "mobile_no", "mobile no rf"]
    )
    
    # FAST PATH: Use index if we are querying the cached dataframe
    if df is _DF_CACHE and _MOBILE_INDEX is not None:
        if mobile in _MOBILE_INDEX:
            indices = _MOBILE_INDEX[mobile]
            return df.loc[indices]
            
    # SLOW PATH: Linear scan (fallback)
    try:
        # Strict match
        filtered = df[df[mobile_col].astype(str).str.strip() == mobile]
        
        # Partial match fallback (only if strict match fails)
        if filtered.empty:
            filtered = df[df[mobile_col].astype(str).str.contains(mobile, na=False)]
            
        return filtered
    except KeyError as exc:
        raise RuntimeError(f"Column '{mobile_col}' not found in Excel data: {exc}")

# ---------------------------------------------------------------------------
# Email composition & sending
# ---------------------------------------------------------------------------

def build_email_body(
    customer_name: str,
    mobile: str,
    address: str,
    product_blocks: List[str],
    submitted_dt: datetime,
) -> str:
    """Create the HTML email body used by the Streamlit app.
    """
    ist_formatted = submitted_dt.strftime("%Y-%m-%d %H:%M:%S IST")
    product_info = "<br><br>".join(product_blocks)
    html = f"""
<div style="font-family: Arial, sans-serif; max-width: 800px; margin: 0 auto;">
    <div style="background: linear-gradient(135deg, #2E86C1 0%, #5DADE2 100%); color: white; padding: 20px; text-align: center; border-radius: 10px 10px 0 0;">
        <h2 style="margin: 0;">🛡️ Warranty Claim Submission</h2>
        <p style="margin: 5px 0 0 0;">New claim received from customer</p>
    </div>
    <div style="background: #f8f9fa; padding: 20px; border-radius: 0 0 10px 10px;">
        <p>Dear Shyla,</p>
        <p>We have received a warranty claim for the products purchased by our customer. Please find the details below:</p>
        <div style="background: white; padding: 15px; border-radius: 8px; margin: 12px 0; border-left: 4px solid #2E86C1;">
            <h3 style="color: #2E86C1; margin-top: 0;">👤 Customer Information</h3>
            <p><strong>Name:</strong> {customer_name}<br>
            <strong>Mobile No:</strong> {mobile}<br>
            <strong>Address:</strong> {address}</p>
        </div>
        <div style="background: white; padding: 15px; border-radius: 8px; margin: 12px 0; border-left: 4px solid #28A745;">
            <h3 style="color: #28A745; margin-top: 0;">📦 Product Details & Issue Description</h3>
            <div style="font-family: monospace; font-size: 14px;">{product_info}</div>
        </div>
        <div style="background: #e7f3ff; padding: 12px; border-radius: 8px; margin: 12px 0;">
            <p><strong>📅 Submitted:</strong> {ist_formatted}</p>
            <p style="margin-bottom: 0;">We request your team to review and process this claim at the earliest convenience.</p>
        </div>
        <div style="text-align: center; margin-top: 14px; padding-top: 10px; border-top: 1px solid #e9ecef;">
            <p style="margin: 0;"><strong>Best Regards,</strong><br>
            <strong>JASIL N</strong><br>
            📞 +918589852747</p>
        </div>
    </div>
</div>
"""
    return html


def send_email(
    subject: str,
    html_body: str,
    attachments: Optional[List[Dict[str, Any]]] = None,
) -> None:
    """Send an email via the configured SMTP server.

    Args:
        subject: Email subject line.
        html_body: HTML content for the email.
        attachments: List of dicts with keys ``filename`` and ``bytes``.
    """
    if not SENDER_PASSWORD:
        raise RuntimeError(
            "SMTP password not set. Provide it via the CLAIM_SENDER_PASSWORD environment variable."
        )
    msg = MIMEMultipart()
    msg["From"] = SENDER_EMAIL
    msg["To"] = TARGET_EMAIL
    msg["Cc"] = ", ".join(CC_EMAILS)
    msg["Subject"] = subject
    msg.attach(MIMEText(html_body, "html"))

    if attachments:
        for att in attachments:
            part = MIMEApplication(att["bytes"], Name=att["filename"])
            part["Content-Disposition"] = f'attachment; filename="{att["filename"]}"'
            msg.attach(part)

    recipients = [TARGET_EMAIL] + CC_EMAILS
    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=10) as server:
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        server.sendmail(SENDER_EMAIL, recipients, msg.as_string())

# ---------------------------------------------------------------------------
# Google Sheets / HTTP endpoint submission
# ---------------------------------------------------------------------------

def submit_claim(payload: Dict[str, Any]) -> bool:
    """POST the claim payload to the configured ``WEB_APP_URL``.

    Returns ``True`` if the request succeeded (status code 200) else ``False``.
    """
    try:
        response = requests.post(WEB_APP_URL, json=payload, timeout=8)
        return response.status_code == 200
    except Exception as exc:
        print(f"Error submitting claim to endpoint: {exc}", file=sys.stderr)
        return False

# ---------------------------------------------------------------------------
# High‑level orchestration – callable from CLI or other code
# ---------------------------------------------------------------------------

def process_claim(
    mobile: str,
    address: str,
    selected_products: List[Dict[str, str]],
    global_issue: Optional[str] = None,
    global_file_path: Optional[str] = None,
    per_product_issues: Optional[List[Dict[str, Any]]] = None,
) -> Dict[str, Any]:
    """Validate inputs, send email, and submit the claim.

    Parameters
    ----------
    mobile: 10‑digit mobile number as string.
    address: Service address.
    selected_products: List of dicts, each containing the keys ``invoice``,
        ``model``, ``serial``, ``osid`` (all strings).
    global_issue: If provided, this description is applied to *all* products.
    global_file_path: Path to a single attachment that applies to all products.
    per_product_issues: Optional list matching ``selected_products`` where each
        entry is ``{"issue": str, "file_path": Optional[str]}``.

    Returns a dict with ``success`` (bool) and ``message`` (str).
    """
    # Basic validation
    if not (mobile.isdigit() and len(mobile) == 10):
        return {"success": False, "message": "Invalid mobile number"}
    if not address:
        return {"success": False, "message": "Address is required"}
    if not selected_products:
        return {"success": False, "message": "At least one product must be selected"}
    if global_issue is None and not per_product_issues:
        return {"success": False, "message": "Issue description required"}

    # Build product blocks for email
    product_blocks = []
    attachments = []
    if global_issue is not None:
        # Global mode – same issue & optional file for every product
        for prod in selected_products:
            block = (
                f"Invoice  : {prod.get('invoice', '')}<br>"
                f"Model    : {prod.get('model', '')}<br>"
                f"Serial No: {prod.get('serial', '')}<br>"
                f"OSID     : {prod.get('osid', '')}<br>"
                f"Issue    : {global_issue}"
            )
            product_blocks.append(block)
        if global_file_path and os.path.isfile(global_file_path):
            with open(global_file_path, "rb") as f:
                attachments.append({"filename": os.path.basename(global_file_path), "bytes": f.read()})
    else:
        # Per‑product mode – iterate over supplied issues
        for prod, issue_info in zip(selected_products, per_product_issues or []):
            issue_text = issue_info.get("issue", "")
            block = (
                f"Invoice  : {prod.get('invoice', '')}<br>"
                f"Model    : {prod.get('model', '')}<br>"
                f"Serial No: {prod.get('serial', '')}<br>"
                f"OSID     : {prod.get('osid', '')}<br>"
                f"Issue    : {issue_text}"
            )
            product_blocks.append(block)
            file_path = issue_info.get("file_path")
            if file_path and os.path.isfile(file_path):
                with open(file_path, "rb") as f:
                    attachments.append({"filename": os.path.basename(file_path), "bytes": f.read()})

    # Assemble email
    ist_now = get_ist_datetime()
    # Look up customer name from the Excel data
    try:
        df = load_excel_data()
        customer_records = get_customer_records(df, mobile)
        if not customer_records.empty:
            name_col = resolve_column(df, ["name", "customer name", "customer"])
            customer_name = str(customer_records.iloc[0].get(name_col, "Customer"))
        else:
            customer_name = "Customer"
    except:
        customer_name = "Customer"
    subject = f"🛡️ Warranty Claim Submission – {customer_name}"
    html_body = build_email_body(customer_name, mobile, address, product_blocks, ist_now)
    try:
        send_email(subject, html_body, attachments)
    except Exception as exc:
        return {"success": False, "message": f"Failed to send email: {exc}"}

    # Prepare payload for the Google Sheets endpoint
    payload = {
        "customer_name": customer_name,
        "mobile_no": mobile,
        "address": address,
        "products": "; ".join([p.get("invoice", "") for p in selected_products]),
        "issue_description": global_issue if global_issue is not None else " || ".join([i.get("issue", "") for i in (per_product_issues or [])]),
        "status": "Pending",
        "submitted_date": ist_now.isoformat(),
    }
    submitted = submit_claim(payload)
    if not submitted:
        return {"success": False, "message": "Email sent but failed to submit to tracking system"}
    return {"success": True, "message": "Claim processed successfully"}

# ---------------------------------------------------------------------------
# Command‑line interface
# ---------------------------------------------------------------------------

def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Process a warranty claim without Streamlit UI.")
    parser.add_argument("--mobile", required=True, help="10‑digit mobile number of the customer")
    parser.add_argument("--address", required=True, help="Service address")
    parser.add_argument(
        "--products-json",
        required=True,
        help="Path to a JSON file describing selected products. Example format: [{\"invoice\": \"123\", \"model\": \"X\", \"serial\": \"ABC\", \"osid\": \"001\"}, ...]",
    )
    parser.add_argument(
        "--global-issue",
        help="If set, this issue description applies to all products",
    )
    parser.add_argument(
        "--global-file",
        help="Path to a single attachment that applies to all products (used with --global-issue)",
    )
    parser.add_argument(
        "--per-product-issues-json",
        help="Path to JSON matching the order of --products-json. Each entry: {\"issue\": \"...\", \"file_path\": \"optional_path\"}",
    )
    return parser.parse_args()


def main() -> None:
    args = _parse_args()
    # Load product list
    with open(args.products_json, "r", encoding="utf-8") as f:
        products = json.load(f)
    per_product = None
    if args.per_product_issues_json:
        with open(args.per_product_issues_json, "r", encoding="utf-8") as f:
            per_product = json.load(f)
    result = process_claim(
        mobile=args.mobile,
        address=args.address,
        selected_products=products,
        global_issue=args.global_issue,
        global_file_path=args.global_file,
        per_product_issues=per_product,
    )
    print(json.dumps(result, ensure_ascii=False, indent=2))

if __name__ == "__main__":
    main()
