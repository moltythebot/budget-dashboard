"""Import FNB Premier Current Account statements into budget database.

This imports transactions from 5 FNB PDF statements covering:
- 225: 28 Feb 2026 - 31 Mar 2026
- 226: 31 Mar 2026 - 30 Apr 2026
- 227: 30 Apr 2026 - 30 May 2026
- 228: 30 May 2026 - 30 Jun 2026
- 229: 30 Jun 2026 - 31 Jul 2026

Classification rules (matching import_personal_budget.py):
- Income: Salary deposits (FNB OB Pmt Salary ...)
- Savings: R1,400 Monthly Transfer (if present)
- Expenses (kept):
  - PSG AML R8,500/month as "Retirement (PSG Aml)"
  - Sanlamgap as insurance
  - Bank fees
  - Personal recurring (tracker, ooba, netstar, titanium, vodacom, multid, afrihost)
  - Electricity, rates, car license
  - Garden services, cleaning
  - Personal meals, travel, subscriptions, medical
- Excluded (business/transfers/investment):
  - DJI/Mini/Mavic/Avata/Neo/Air/FPV/drone purchases (Gerhard...)
  - Credit card clearing transfers
  - Investment/savings transfers
  - Pinegrove, Coenie Betaaling, Huis Betaal
  - Tesla share, investment-profit
  - Maintenance above R4,500
  - SARS
  - Stock/courier/DJI/FPV/Air3S/Aliexpress/Bob Group
"""

from __future__ import annotations

from datetime import datetime
import json
import re
import sqlite3
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
DB_PATH = BASE_DIR / "budget.db"
TRANSACTIONS_JSON = Path("/home/molty/transactions.json")
EXCEL_PATH = Path("/home/molty/.hermes/cache/documents/doc_1113d65cf0cc_2026_March_-_2026_Aug-_Provisional_tax_info_sheet_-_SR.xlsx")

CAT_FIX = {
    "insurnance": "insurance",
    "insurnace": "insurance",
    "maintennace": "maintenance",
    "accountng fees": "accounting fees",
}

DISPLAY_CATEGORY = {
    "accounting fees": "accounting fees",
    "bank fees": "bank fees",
    "cleaning": "cleaning",
    "donation": "donation",
    "electricity": "electricity",
    "entertainment": "entertainment",
    "family": "family",
    "insurance": "insurance",
    "internet": "internet",
    "maintenance": "maintenance",
    "meals": "meals",
    "medical": "medical",
    "rates": "rates",
    "security": "security",
    "subscription": "subscription",
    "telephone": "telephone",
    "travel": "travel",
    "uif": "uif",
}

BUSINESS_CATEGORIES = {"stock", "courier"}
BUSINESS_KEYWORDS = [
    "stock", "courier", "fpv", "dji", "air3s", "aliexpress", "bob group",
    "mini", "mavic", "avata", "neo", "air 3", "air2s", "rc pro", "goggles",
    "battery", "lito", "matrice", "drone", "fpv", "controller", "charger",
    "smart controller", "m2ea", "m30t", "insta 360", "ace pro", "air 1",
    "mavic air", "osmo pocket", "mini 2", "mini 3", "mini 4", "mini 5",
    "avata 2", "avata 3", "neo 2", "air 2s", "air 3", "rcn1", "rc n1",
    "goggles 2", "goggles 3", "goggles integra", "neo case", "flymore",
    "props", "mavic 2", "mavic 3", "mavic 3t", "lito x1", "mini 4k",
    "mini 3 pro", "mini 4 pro", "mini 5 pro", "neo motion", "advanced",
    "enterprise", "dji enterprise", "antigravity", "air 3s", "neo 2 kit",
    "drone stock", "dji drones", "dji mini", "dji mavic", "dji avata",
    "dji neo", "dji air", "dji fpv", "dji rc", "dji goggles",
    "mini 3 fmc", "mini 2 fmc", "mini 5", "mini 4+", "mini 3+",
    "air 3 extra", "air 3 flymore", "air 2", "m2ea", "m2e",
    "go pro", "gopro", "insta 360", "laptop", "gaming tower", "gaming pc",
    "camera", "mavic 3t", "mavic 3 enterprise", "m30t", "m30",
    "trade in", "5x chargers"
]
TRANSFER_KEYWORDS = [
    "transfer", "scheduled trf", "scheduled transfer", "internal transfer",
    "monthly transfer", "clear credit card", "investment", "savings",
    "loan repay", "load repay"
]

EXCLUDE_KEYWORDS = [
    "pinegrove", "coenie betaaling", "huis betaal", "tesla share",
    "investment profit", "investment-profit", "stock", "courier"
]

def clean_text(value) -> str:
    if value is None:
        return ""
    return " ".join(str(value).strip().split())

def parse_statement_date(date_str: str, statement_period: str) -> datetime:
    """Parse 'DD Mon' format to datetime using statement period for year."""
    # Check if already ISO format (YYYY-MM-DD)
    if re.match(r'^\d{4}-\d{2}-\d{2}$', date_str):
        return datetime.strptime(date_str, "%Y-%m-%d")
    
    # Extract year from period (e.g., "28 February 2026 to 31 March 2026")
    year_match = re.search(r'20\d{2}', statement_period)
    base_year = int(year_match.group()) if year_match else 2026
    
    try:
        dt = datetime.strptime(f"{date_str} {base_year}", "%d %b %Y")
    except ValueError:
        # Try next year if month is Jan-Feb and period starts late in year
        dt = datetime.strptime(f"{date_str} {base_year + 1}", "%d %b %Y")
    
    return dt

def month_key(dt: datetime) -> str:
    return dt.strftime("%Y-%m")

def is_salary_deposit(desc: str) -> bool:
    dl = desc.lower()
    return "salary" in dl and "ob pmt" in dl

 # Human-readable glosses so the dashboard's Notes column explains each item.
NOTE_RULES = [
    (("tracker",), "Vehicle tracking debit order (Tracker)"),
    (("netstar",), "Vehicle tracking debit order (Netstar)"),
    (("titanium",), "Vehicle tracking debit order (Titanium)"),
    (("sanlamgap",), "Vehicle gap-cover insurance debit order (Sanlam)"),
    (("internal debit order fnb st ins",), "FNB insurance debit order"),
    (("ooba",), "Building insurance debit order"),
    (("insurance",), "Insurance debit order"),
    (("sanlam",), "Sanlam insurance premium"),
    (("vodacom",), "Cellphone contract (Vodacom)"),
    (("afrihost",), "Fibre internet (Afrihost)"),
    (("multid",), "Charity donation debit order"),
    (("formealswheel",), "Charity donation debit order"),
    (("durbanville kleuterakad",), "SKA KleuterAkademie school fees"),
    (("uif",), "UIF statutory contribution"),
    (("payroll",), "Payroll statutory payment"),
    (("electric",), "Electricity purchase"),
    (("prepaid",), "Prepaid electricity purchase"),
    (("pos purchase easypay",), "Municipal rates via Easypay"),
    (("a coct",), "City of Cape Town municipal account"),
    (("municipal",), "Municipal account payment"),
    (("car license",), "Car licence renewal"),
    (("lic renewal",), "Licence renewal"),
    (("traffic fine",), "Traffic fine"),
    (("garden",), "Garden service"),
    (("cleaning",), "Cleaning service / domestic wages"),
    (("jocelyn",), "Domestic wages (Jocelyn)"),
    (("aircon",), "Aircon repair/service"),
    (("fridge",), "Fridge repair/purchase"),
    (("battery centre",), "Battery replacement"),
    (("bin",), "Bin cleaning service"),
    (("accounting",), "Accounting fees"),
    (("jolandi betaal",), "Medical payment to Jolandi"),
    (("willem olivier",), "Family support payment"),
    (("deposit for housefix",), "House fix deposit"),
    (("payments for rest",), "Rest-day payment"),
    (("verjaarsdag",), "Birthday entertainment"),
    (("apie",), "Entertainment outing"),
    (("flowers",), "Flowers gift"),
    (("netflix",), "Netflix subscription"),
    (("google one",), "Google One storage"),
    (("google  one",), "Google One storage"),
    (("takealot",), "Takealot online order"),
    (("wondershare",), "Wondershare software licence"),
    (("suno",), "Suno AI subscription"),
    (("xtra savings",), "Checkers Xtra Savings membership"),
    (("melon mobile",), "Melon Mobile subscription"),
    (("tv subscription",), "TV subscription"),
    (("subscription",), "Subscription service"),
    (("dischem",), "Pharmacy (Dischem)"),
    (("clicks",), "Pharmacy (Clicks)"),
    (("pharmacy",), "Pharmacy purchase"),
    (("engen",), "Fuel (Engen)"),
    (("shell",), "Fuel (Shell)"),
    (("uber",), "Uber trip"),
    (("huguenot",), "Toll fee (Huguenot Tunnel)"),
    (("tunnel",), "Toll fee"),
    (("food",), "Meal purchase"),
    (("lunch",), "Lunch"),
    (("biltong",), "Biltong snack"),
]

CATEGORY_FALLBACK_NOTES = {
    "meals": "Grocery/dining spend",
    "travel": "Travel/fuel spend",
    "subscription": "Subscription service",
    "medical": "Medical/pharmacy spend",
    "maintenance": "Home maintenance",
    "bank fees": "Bank charge",
    "insurance": "Insurance premium",
    "security": "Security service",
    "telephone": "Cellphone airtime/data",
    "internet": "Internet service",
    "donation": "Donation",
    "electricity": "Electricity",
    "rates": "Rates & utilities",
    "cleaning": "Cleaning/wages",
    "uif": "UIF contribution",
    "entertainment": "Entertainment spend",
    "family": "Family support",
    "accounting fees": "Accounting fees",
}


def note_for(desc_lower: str, display_category: str) -> str:
    for keys, phrase in NOTE_RULES:
        if any(k in desc_lower for k in keys):
            return phrase
    return CATEGORY_FALLBACK_NOTES.get(
        str(display_category).lower(),
        "Personal expense - see description."
    )


def classify_transaction(desc: str, amount: float) -> tuple[str, str, str, str] | None:
    """Classify a transaction. Returns (entry_type, category, subcategory, notes) or None to exclude."""
    dl = desc.lower()

    # All non-salary inflows are business money (Gerhard's ruling): salary is
    # handled upstream in build_entries(), so any other credit stays out of the
    # personal budget - stock deposits (Uys Burger), interest rebates, RTC
    # refund credits, drone-sale receipts, all of them.
    if amount > 0:
        return None
    
    # Monthly 1K Transfer - savings
    if "scheduled trf to monthly 1k transfer" in dl or ("monthly" in dl and "1k" in dl and "transfer" in dl):
        return (
            "savings",
            "Monthly Transfer Savings",
            desc,
            "R1,400 monthly scheduled transfer; recorded as savings per Gerhard's correction."
        )
    
    # PSG AML - retirement expense
    if "psg aml" in dl:
        return (
            "expense",
            "Retirement (PSG Aml)",
            desc,
            "Retirement annuity debit order; kept as spending/outflow per Gerhard's instruction."
        )
    
    # Bank fees (blank or short descriptions with small amounts, or explicit bank fee)
    if (len(desc.strip()) <= 3 or desc.strip() == "") and amount < 0:
        # All blank/short description debits are bank fees
        return (
            "expense",
            "bank fees",
            "Bank fees",
            "Bank fee (blank description)."
        )
    # Common bank fee amounts (based on existing data patterns)
    if desc.strip() == "" and amount < 0 and abs(amount) in [2.00, 10.00, 20.00, 30.00, 50.00, 69.00, 80.00, 90.00, 100.00, 110.00, 275.00, 285.00]:
        return (
            "expense",
            "bank fees",
            "Bank fees",
            "Bank fee (known amount pattern)."
        )
    if "int on debit balance" in dl or "dr interest rebate" in dl:
        return (
            "expense",
            "bank fees",
            desc,
            "Interest/rebate on debit balance."
        )
    if "int pymt fee" in dl or "#int pymt" in dl:
        return (
            "expense",
            "bank fees",
            desc,
            "International payment fee."
        )
    # POS Purchase Easypay City - rates/municipal
    if "pos purchase easypay" in dl and "city" in dl:
        return (
            "expense",
            "rates",
            desc,
            "Municipal payment via Easypay."
        )
    
    # Business drone purchases - EXCLUDE
    if any(kw in dl for kw in BUSINESS_KEYWORDS):
        return None
    
    # Credit card clearing - EXCLUDE
    if "clear credit card" in dl:
        return None
    
    # Investment/savings/loan transfers - EXCLUDE
    if any(kw in dl for kw in TRANSFER_KEYWORDS):
        return None
    
    # Other exclusions
    if any(kw in dl for kw in EXCLUDE_KEYWORDS):
        return None
    # Aloha Holdings - business (Infinity Drones)
    elif "aloha holdings" in dl:
        return None  # Business stock purchase
    # Mohamed Badrudin - business stock
    elif "mohamed badrudin" in dl:
        return None  # Business stock purchase
    # Janse Van Rensbur - business stock
    elif "janse van rensbur" in dl:
        return None  # Business stock purchase
    # Rtc Credit Wa Burger Kaalplek - business stock
    elif "wa burger kaalplek" in dl:
        return None  # Business stock purchase
    # Maseeg Damon - business stock (Infinity Drones)
    elif "maseeg damon" in dl:
        return None  # Business stock purchase
    
    # Maintenance above R4,500 - EXCLUDE
    if "maintenance" in dl and abs(amount) > 4500:
        return None
    
    # SARS - EXCLUDE
    if "sars" in dl:
        return None
    
    # Recurring personal expenses - categorize
    category = ""
    if "tracker" in dl:
        category = "security"
    elif "ooba" in dl or "building" in dl:
        category = "insurance"
    elif "netstar" in dl:
        category = "security"
    elif "titanium" in dl:
        category = "security"
    elif "vodacom" in dl:
        category = "telephone"
    elif "afrihost" in dl:
        category = "internet"
    elif "multid" in dl or "formealswheel" in dl:
        category = "donation"
    elif "sanlamgap" in dl:
        category = "insurance"
    elif "uif" in dl:
        category = "uif"
    elif "electric" in dl or "prepaid" in dl:
        category = "electricity"
    elif "rate" in dl or "municipal" in dl or "a coct" in dl:
        category = "rates"
    elif "license" in dl or "lic renewal" in dl:
        category = "rates"
    elif "garden" in dl:
        category = "maintenance"
    elif "repair" in dl and "inv" in dl and abs(amount) <= 4500:
        category = "maintenance"
    elif "cleaning" in dl or "jocelyn" in dl:
        category = "cleaning"
    elif "payroll" in dl:
        category = "uif"
    elif "traffic fine" in dl:
        category = "travel"
    elif "car license" in dl or "lic renewal" in dl:
        category = "rates"
    elif "insurance" in dl or "sanlam" in dl:
        category = "insurance"
    elif "accounting" in dl:
        category = "accounting fees"
    elif "food" in dl or "lunch" in dl or "biltong" in dl:
        category = "meals"
    elif "spar" in dl or "checkers" in dl or "pick n pay" in dl or "woolworths" in dl or "kfc" in dl or "mcd" in dl or "pizza" in dl or "burger" in dl or "restaurant" in dl or "cafe" in dl or "coffee" in dl or "bamboo garden" in dl or "spier" in dl or "guano" in dl or "two oceans" in dl or "bootlegger" in dl or "kauai" in dl or "safari" in dl or "yoco" in dl or "vida" in dl or "takoda" in dl or "de oude" in dl or "starke" in dl or "pnp" in dl or "atkv" in dl or "senqu" in dl:
        category = "meals"
    elif "engen" in dl or "shell" in dl or "uber" in dl or "huguenot" in dl or "tunnel" in dl or "travel" in dl or "drone trip" in dl:
        category = "travel"
    elif "netflix" in dl or "google one" in dl or "google  one" in dl or "takealot" in dl or "wondershare" in dl or "suno" in dl or "xtra savings" in dl or "subscription" in dl or "tv subscription" in dl or "durbanville kleuterakad" in dl or "melon mobile" in dl or "payfast" in dl:
        category = "subscription"
    elif "medical" in dl or "dischem" in dl or "clicks" in dl or "pharmacy" in dl:
        category = "medical"
    elif "courier" in dl and "net" in dl:
        category = "travel"
    elif "verjaarsdag" in dl or "apie" in dl:
        category = "entertainment"
    elif "flowers" in dl:
        category = "entertainment"
    elif "bin" in dl and "clean" in dl:
        category = "maintenance"
    elif "mowers" in dl or "mica" in dl or "bwh" in dl or "hifi" in dl or "csc" in dl or "battery centre" in dl:
        category = "maintenance"
    elif "aircon" in dl:
        category = "maintenance"
    elif "fridge" in dl:
        category = "maintenance"
    elif "spark" in dl and "payshap" in dl:
        # This is business - Spark controller commission
        return None
    # Personal/family transactions
    elif "deposit for housefix" in dl:
        category = "maintenance"
    elif "payments for rest" in dl and "gerhard" in dl:
        category = "entertainment"
    elif "jolandi betaal" in dl:
        category = "medical"
    elif "willem olivier" in dl:
        category = "family"
    elif "refund" in dl and "gerhard" in dl:
        category = "Uncategorised"  # Refund given by Gerhard - personal expense
    elif "send money app" in dl:
        category = "family"  # Personal transfer
    elif "payshap account off-us" in dl and ("huawei" in dl or "p4p" in dl):
        return None  # Personal electronics/unknown - exclude as business stock
    # (Uys Burger / interest rebates / RTC credits: handled by the positive-
    #  amount guard at the top - inflows never reach the categorisation ladder)
    # FNB ST Insurance - internal debit order
    elif "internal debit order fnb st ins" in dl:
        category = "insurance"
    # Maurice commission - business drone
    elif "maurice commission" in dl:
        return None  # Business expense, exclude
    # TB48 batteries - business
    elif "tb48" in dl or "tb 48" in dl:
        return None  # Business drone battery
    # Rtc Credit Wa Burger Kaalplek - business stock
    elif "wa burger kaalplek" in dl:
        return None  # Business stock purchase
    # Maseeg Damon - business stock (Infinity Drones)
    elif "maseeg damon" in dl:
        return None  # Business stock purchase
    
    if category:
        display = DISPLAY_CATEGORY.get(category, category or "Uncategorised")
        return (
            "expense",
            display,
            desc,
            note_for(dl, display)
        )
    
    # Default: categorize as uncategorised expense if negative amount
    if amount < 0:
        return (
            "expense",
            "Uncategorised",
            desc,
            "Uncategorised personal debit - needs manual review."
        )
    
    # Positive amounts that aren't salary - likely business income, exclude
    return None

def load_transactions() -> list[dict]:
    with open(TRANSACTIONS_JSON) as f:
        return json.load(f)


def _load_fee_descriptions() -> dict:
    """(date, amount) -> [descriptions] from the spreadsheet's cheque-debit block."""
    from collections import defaultdict
    out = defaultdict(list)
    try:
        import openpyxl
        wb = openpyxl.load_workbook(EXCEL_PATH, read_only=True, data_only=True)
        ws = wb["March-Aug 2026"]
        for r in ws.iter_rows(values_only=True):
            if len(r) < 8 or r[5] is None or r[7] is None:
                continue
            desc = str(r[6]).strip() if r[6] is not None else ""
            if desc:
                out[(str(r[5])[:10], round(float(r[7]), 2))].append(desc)
        wb.close()
    except Exception as exc:  # spreadsheet unavailable -> keep generic labels
        print(f"Note: fee-description recovery skipped ({exc})")
    return out

def build_entries(transactions: list[dict]) -> list[tuple]:
    entries = []
    fee_names = _load_fee_descriptions()
    
    for t in transactions:
        dt = parse_statement_date(t['date'], t['period'])
        mk = month_key(dt)
        desc = t['description']
        amount = abs(t['amount'])  # Use absolute value, entry_type determines direction
        
        # Blank statement descriptions (bank charges): recover the real wording
        # from the provisional-tax spreadsheet so Detail isn't just "Bank fees".
        if (desc is None or str(desc).strip() == "") and t["amount"] < 0:
            pool = fee_names.get((t["date"][:10], round(abs(t["amount"]), 2)))
            if pool:
                recovered = pool.pop(0)
                entries.append((
                    mk,
                    "bank fees",
                    recovered,
                    amount,
                    "expense",
                    "Bank charge - wording recovered from provisional-tax spreadsheet.",
                    "fnb_statement",
                    0,
                ))
                continue
        
        if is_salary_deposit(desc):
            entries.append((
                mk,
                "Cheque Credit / Net Salary",
                desc,
                amount,
                "income",
                "Actual salary paid into cheque credit account; business income ignored.",
                "fnb_statement",
                0,
            ))
        else:
            classified = classify_transaction(desc, t['amount'])
            if classified is None:
                continue
            entry_type, category, subcategory, notes = classified
            entries.append((
                mk,
                category,
                subcategory,
                amount,
                entry_type,
                notes,
                "fnb_statement",
                0,
            ))
    
    return entries

def reset_fnb_imported_rows(conn: sqlite3.Connection):
    conn.execute("DELETE FROM budget_entries WHERE source='fnb_statement'")
    # Don't delete monthly_snapshots - we'll rebuild them

def insert_entries(conn: sqlite3.Connection, entries: list[tuple]):
    conn.executemany(
        """
        INSERT INTO budget_entries
            (month, category, subcategory, amount, entry_type, notes, source, is_sample)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """,
        entries,
    )

def rebuild_snapshots(conn: sqlite3.Connection):
    rows = conn.execute(
        """
        SELECT month,
               SUM(CASE WHEN entry_type='income' THEN amount ELSE 0 END) AS income,
               SUM(CASE WHEN entry_type='expense' THEN amount ELSE 0 END) AS expenses,
               SUM(CASE WHEN entry_type='savings' THEN amount ELSE 0 END) AS savings
        FROM budget_entries
        GROUP BY month
        ORDER BY month
        """
    ).fetchall()

    for month, income, expenses, savings in rows:
        income = income or 0
        expenses = expenses or 0
        savings = savings or 0
        conn.execute(
            """
            INSERT INTO monthly_snapshots
                (month, total_income, total_expenses, total_savings, net_cashflow)
            VALUES (?, ?, ?, ?, ?)
            ON CONFLICT(month) DO UPDATE SET
                total_income=excluded.total_income,
                total_expenses=excluded.total_expenses,
                total_savings=excluded.total_savings,
                net_cashflow=excluded.net_cashflow
            """,
            (month, income, expenses, savings, income - expenses - savings),
        )

def print_verification(conn: sqlite3.Connection):
    print("FNB Imported rows:", conn.execute("SELECT COUNT(*) FROM budget_entries WHERE source='fnb_statement'").fetchone()[0])
    print("Sample rows:", conn.execute("SELECT COUNT(*) FROM budget_entries WHERE is_sample=1").fetchone()[0])
    print("By type:")
    for row in conn.execute("SELECT entry_type, COUNT(*), ROUND(SUM(amount), 2) FROM budget_entries WHERE source='fnb_statement' GROUP BY entry_type ORDER BY entry_type"):
        print(f"  {row[0]}: rows={row[1]}, total=R{row[2]:,.2f}")

    print("\nMonthly summary (FNB only):")
    rows = conn.execute(
        """
        SELECT month,
               ROUND(SUM(CASE WHEN entry_type='income' THEN amount ELSE 0 END), 2) AS income,
               ROUND(SUM(CASE WHEN entry_type='expense' THEN amount ELSE 0 END), 2) AS expenses,
               ROUND(SUM(CASE WHEN entry_type='savings' THEN amount ELSE 0 END), 2) AS savings,
               ROUND(SUM(CASE WHEN entry_type='income' THEN amount ELSE 0 END)
                 - SUM(CASE WHEN entry_type='expense' THEN amount ELSE 0 END)
                 - SUM(CASE WHEN entry_type='savings' THEN amount ELSE 0 END), 2) AS net
        FROM budget_entries
        WHERE source='fnb_statement'
        GROUP BY month
        ORDER BY month
        """
    ).fetchall()
    for month, income, expenses, savings, net in rows:
        print(f"  {month}: income=R{income:,.2f}, expenses=R{expenses:,.2f}, savings=R{savings:,.2f}, net=R{net:,.2f}")

    print("\nExpense categories (FNB):")
    for category, total, count in conn.execute("SELECT category, ROUND(SUM(amount), 2), COUNT(*) FROM budget_entries WHERE source='fnb_statement' AND entry_type='expense' GROUP BY category ORDER BY SUM(amount) DESC"):
        print(f"  {category}: R{total:,.2f} ({count})")

def main():
    if not TRANSACTIONS_JSON.exists():
        raise SystemExit(f"Transactions file not found: {TRANSACTIONS_JSON}")

    transactions = load_transactions()
    entries = build_entries(transactions)

    conn = sqlite3.connect(DB_PATH)
    try:
        reset_fnb_imported_rows(conn)
        insert_entries(conn, entries)
        rebuild_snapshots(conn)
        conn.commit()
        print_verification(conn)
    finally:
        conn.close()

if __name__ == "__main__":
    main()