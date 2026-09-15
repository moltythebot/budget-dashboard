#!/usr/bin/env python3
"""Import FNB Credit Card transactions from Excel into budget database."""

from __future__ import annotations

from datetime import datetime
import json
import sqlite3
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
DB_PATH = BASE_DIR / "budget.db"
CC_TRANSACTIONS_JSON = Path("/home/molty/credit_card_transactions.json")

# Display categories (same as import_fnb_statements.py)
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

# Business categories to exclude
BUSINESS_CATEGORIES = {"stock", "courier"}

# Business keywords (same as FNB import)
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
    "monthly transfer", "clear credit card", "investment",
    "loan repay", "load repay", "savings transfer", "transfer to savings"
]

EXCLUDE_KEYWORDS = [
    "pinegrove", "coenie betaaling", "huis betaal", "tesla share",
    "investment profit", "investment-profit", "stock", "courier"
]


def month_key(dt: datetime) -> str:
    return dt.strftime("%Y-%m")


def net_fully_reversed(transactions: list[dict]) -> list[dict]:
    """Remove charge/reversal pairs that cancel exactly.

    A positive charge is dropped only when a negative line dated the same day,
    for the same absolute amount, has 'reversal' in its description (e.g. Slow
    Lounge Entry Fee + Slow Entry Fee Reversal). Matched charges are dropped,
    and reversal lines are always dropped here too - not left to the
    classifier's amount<0 guard. Unpaired charges are kept; other negative
    lines still fall through to that guard.

    Limitation: pairing keys on (date, abs(amount)) only, not description. If a
    future day holds two DIFFERENT same-amount charges plus one reversal, one
    legitimate charge would be dropped. Safe with current data (verified: the
    only +-R300 lines on 2026-03-20 are the fee/reversal pairs themselves);
    tighten to description-root matching if that pattern ever appears.
    """
    from collections import Counter

    open_reversals: Counter = Counter()
    for t in transactions:
        amt = t.get("amount") or 0
        desc = (t.get("description") or "").lower()
        if amt < 0 and "reversal" in desc:
            open_reversals[(t.get("date"), round(abs(amt), 2))] += 1

    kept: list[dict] = []
    for t in transactions:
        amt = t.get("amount") or 0
        desc = (t.get("description") or "").lower()
        key = (t.get("date"), round(abs(amt), 2))
        if amt > 0 and "reversal" not in desc and open_reversals.get(key, 0) > 0:
            open_reversals[key] -= 1
            continue
        # Drop reversal lines themselves (never valid expense rows)
        if amt < 0 and "reversal" in desc:
            continue
        kept.append(t)
    return kept


# Human-readable glosses for card spend so Notes explains the item.
CC_NOTE_RULES = [
    (("openrouter",), "OpenRouter AI API credit"),
    (("netflix",), "Netflix subscription"),
    (("google one",), "Google One storage"),
    (("google  one",), "Google One storage"),
    (("takealot",), "Takealot online order"),
    (("temu",), "Temu online order"),
    (("wondershare",), "Wondershare software licence"),
    (("suno",), "Suno AI subscription"),
    (("xtra savings",), "Checkers Xtra Savings membership"),
    (("melon mobile",), "Melon Mobile subscription"),
    (("durbanville kleuterakad",), "SKA KleuterAkademie school fees"),
    (("checkers sixty60",), "Checkers Sixty60 grocery delivery"),
    (("checkers",), "Checkers groceries"),
    (("woolworths",), "Woolworths purchase"),
    (("pick n pay",), "Pick n Pay groceries"),
    (("spar",), "SPAR groceries"),
    (("ok foods",), "OK Foods groceries"),
    (("kfc",), "KFC meal"),
    (("mcd",), "McDonald's meal"),
    (("dischem",), "Dischem pharmacy"),
    (("clicks",), "Clicks pharmacy"),
    (("mkem",), "MKem pharmacy"),
    (("pharmacy",), "Pharmacy purchase"),
    (("engen",), "Fuel (Engen)"),
    (("shell",), "Fuel (Shell)"),
    (("uber",), "Uber trip"),
    (("huguenot",), "Huguenot Tunnel toll"),
    (("city of c",), "City of Cape Town rates (Easypay)"),
    (("slow lounge",), "Airport lounge fee"),
]

_CC_CATEGORY_FALLBACKS = {
    "meals": "Grocery/dining spend",
    "travel": "Travel/fuel spend",
    "subscription": "Subscription service",
    "medical": "Medical/pharmacy spend",
    "maintenance": "Home maintenance",
    "telephone": "Cellphone airtime/data",
    "electricity": "Electricity",
    "rates": "Rates & utilities",
    "bank fees": "Card transaction fee",
}


def cc_note_for(desc_lower: str, display_category: str) -> str:
    for keys, phrase in CC_NOTE_RULES:
        if any(k in desc_lower for k in keys):
            return phrase
    return _CC_CATEGORY_FALLBACKS.get(str(display_category).lower(), "Card spend - see description.")


def classify_cc_transaction(desc: str, amount: float, category: str) -> tuple[str, str, str, str] | None:
    """Classify a credit card transaction. Returns (entry_type, category, subcategory, notes) or None to exclude."""
    dl = desc.lower()
    
    # Skip refunds (negative amounts on credit card = refunds/credits)
    if amount < 0:
        # These are refunds/credits - skip for now or could categorize
        return None
    
    # Business drone purchases - EXCLUDE
    if any(kw in dl for kw in BUSINESS_KEYWORDS):
        return None
    
    # Credit card clearing - EXCLUDE (payment from cheque to credit card)
    if "clear credit card" in dl or "transfer to clear credit" in dl:
        return None
    
    # Investment/savings/loan transfers - EXCLUDE
    if any(kw in dl for kw in TRANSFER_KEYWORDS):
        return None
    
    # Other exclusions
    if any(kw in dl for kw in EXCLUDE_KEYWORDS):
        return None
    
    # Specific business exclusions
    if "aloha holdings" in dl:
        return None
    elif "mohamed badrudin" in dl:
        return None
    elif "janse van rensbur" in dl:
        return None
    elif "wa burger kaalplek" in dl:
        return None
    elif "maurice commission" in dl:
        return None
    elif "tb48" in dl or "tb 48" in dl:
        return None
    elif "spark" in dl and "payshap" in dl:
        return None
    
    # Use the category from Excel if it's a known good category
    # The Excel already has categories assigned
    final_category = category.lower().strip()
    
    # Map Excel categories to our display categories
    category_mapping = {
        "meals": "meals",
        "travel": "travel",
        "subscription": "subscription",
        "stock": "stock",  # Will be excluded
        "courier": "courier",  # Will be excluded
        "bank fees": "bank fees",
        "telephone": "telephone",
        "rates": "rates",
        "electricity": "electricity",
        "medical": "medical",
        "maintenance": "maintenance",
    }
    
    mapped_category = category_mapping.get(final_category, final_category)
    
    # Exclude business categories
    if mapped_category in BUSINESS_CATEGORIES:
        return None
    
    # For uncategorized, try keyword matching
    if mapped_category == "uncategorised" or mapped_category == "":
        # Keyword-based categorization (same as FNB import)
        if "tracker" in dl:
            mapped_category = "security"
        elif "ooba" in dl or "building" in dl:
            mapped_category = "insurance"
        elif "netstar" in dl:
            mapped_category = "security"
        elif "titanium" in dl:
            mapped_category = "security"
        elif "vodacom" in dl:
            mapped_category = "telephone"
        elif "afrihost" in dl:
            mapped_category = "internet"
        elif "multid" in dl or "formealswheel" in dl:
            mapped_category = "donation"
        elif "sanlamgap" in dl:
            mapped_category = "insurance"
        elif "uif" in dl:
            mapped_category = "uif"
        elif "electric" in dl or "prepaid" in dl:
            mapped_category = "electricity"
        elif "rate" in dl or "municipal" in dl or "a coct" in dl:
            mapped_category = "rates"
        elif "license" in dl or "lic renewal" in dl:
            mapped_category = "rates"
        elif "garden" in dl:
            mapped_category = "maintenance"
        elif "repair" in dl and "inv" in dl and abs(amount) <= 4500:
            mapped_category = "maintenance"
        elif "cleaning" in dl or "jocelyn" in dl:
            mapped_category = "cleaning"
        elif "payroll" in dl:
            mapped_category = "uif"
        elif "traffic fine" in dl:
            mapped_category = "travel"
        elif "car license" in dl or "lic renewal" in dl:
            mapped_category = "rates"
        elif "insurance" in dl or "sanlam" in dl:
            mapped_category = "insurance"
        elif "accounting" in dl:
            mapped_category = "accounting fees"
        elif "food" in dl or "lunch" in dl or "biltong" in dl:
            mapped_category = "meals"
        elif "spar" in dl or "checkers" in dl or "pick n pay" in dl or "woolworths" in dl or "kfc" in dl or "mcd" in dl or "pizza" in dl or "burger" in dl or "restaurant" in dl or "cafe" in dl or "coffee" in dl or "bamboo garden" in dl or "spier" in dl or "guano" in dl or "two oceans" in dl or "bootlegger" in dl or "kauai" in dl or "safari" in dl or "yoco" in dl or "vida" in dl or "takoda" in dl or "de oude" in dl or "starke" in dl or "pnp" in dl or "atkv" in dl or "senqu" in dl or "sixty60" in dl or "asap" in dl or "ok foods" in dl:
            mapped_category = "meals"
        elif "engen" in dl or "shell" in dl or "uber" in dl or "huguenot" in dl or "tunnel" in dl or "travel" in dl or "drone trip" in dl:
            mapped_category = "travel"
        elif "netflix" in dl or "google one" in dl or "google  one" in dl or "takealot" in dl or "wondershare" in dl or "suno" in dl or "xtra savings" in dl or "subscription" in dl or "tv subscription" in dl or "durbanville kleuterakad" in dl or "melon mobile" in dl or "payfast" in dl:
            mapped_category = "subscription"
        elif "medical" in dl or "dischem" in dl or "clicks" in dl or "pharmacy" in dl or "mkem" in dl or "uniclinic" in dl:
            mapped_category = "medical"
        elif "courier" in dl and "net" in dl:
            mapped_category = "travel"
        elif "verjaarsdag" in dl or "apie" in dl:
            mapped_category = "entertainment"
        elif "flowers" in dl:
            mapped_category = "entertainment"
        elif "bin" in dl and "clean" in dl:
            mapped_category = "maintenance"
        elif "mowers" in dl or "mica" in dl or "bwh" in dl or "hifi" in dl or "csc" in dl or "battery centre" in dl:
            mapped_category = "maintenance"
        elif "aircon" in dl:
            mapped_category = "maintenance"
        elif "fridge" in dl:
            mapped_category = "maintenance"
        elif "deposit for housefix" in dl:
            mapped_category = "maintenance"
        elif "payments for rest" in dl and "gerhard" in dl:
            mapped_category = "entertainment"
        elif "jolandi betaal" in dl:
            mapped_category = "medical"
        elif "willem olivier" in dl:
            mapped_category = "family"
        elif "refund" in dl and "gerhard" in dl:
            mapped_category = "Uncategorised"
        elif "send money app" in dl:
            mapped_category = "family"
        elif "payshap account off-us" in dl and ("huawei" in dl or "p4p" in dl):
            return None
        elif "internal debit order fnb st ins" in dl:
            mapped_category = "insurance"
    
    display = DISPLAY_CATEGORY.get(mapped_category, mapped_category or "Uncategorised")
    
    return (
        "expense",
        display,
        desc,
        cc_note_for(dl, display)
    )


def load_cc_transactions() -> list[dict]:
    with open(CC_TRANSACTIONS_JSON) as f:
        return json.load(f)


def reset_cc_imported_rows(conn: sqlite3.Connection):
    conn.execute("DELETE FROM budget_entries WHERE source='fnb_credit_card'")
    conn.commit()


def insert_cc_entries(conn: sqlite3.Connection, entries: list[tuple]):
    conn.executemany(
        """
        INSERT INTO budget_entries
            (month, category, subcategory, amount, entry_type, notes, source, is_sample)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """,
        entries,
    )
    conn.commit()


def rebuild_cc_snapshots(conn: sqlite3.Connection):
    """Refresh monthly_snapshots from budget_entries (mirrors app.py and the FNB importer)."""
    rows = conn.execute(
        """
        SELECT month,
               SUM(CASE WHEN entry_type='income' THEN amount ELSE 0 END),
               SUM(CASE WHEN entry_type='expense' THEN amount ELSE 0 END),
               SUM(CASE WHEN entry_type='savings' THEN amount ELSE 0 END)
        FROM budget_entries
        GROUP BY month
        ORDER BY month
        """
    ).fetchall()
    for month, income, expenses, savings in rows:
        income, expenses, savings = income or 0, expenses or 0, savings or 0
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
    conn.commit()


def main():
    print("Loading credit card transactions...")
    transactions = load_cc_transactions()
    transactions = net_fully_reversed(transactions)
    print(f"Loaded {len(transactions)} credit card transactions")
    
    entries = []
    excluded_count = 0
    
    for t in transactions:
        try:
            dt = datetime.fromisoformat(t['date'].replace(' 00:00:00', ''))
        except ValueError:
            try:
                dt = datetime.strptime(t['date'], "%Y-%m-%d")
            except ValueError:
                print(f"Warning: Could not parse date: {t['date']}")
                continue
        
        mk = month_key(dt)
        desc = t['description']
        amount = abs(t['amount'])
        excel_category = t['category']
        
        classified = classify_cc_transaction(desc, t['amount'], excel_category)
        if classified is None:
            excluded_count += 1
            continue
            
        entry_type, category, subcategory, notes = classified
        entries.append((
            mk,
            category,
            subcategory,
            amount,
            entry_type,
            notes,
            "fnb_credit_card",
            0,
        ))
    
    print(f"Prepared {len(entries)} entries for import ({excluded_count} excluded)")
    
    # Summary by month and category
    from collections import defaultdict
    month_totals = defaultdict(float)
    cat_totals = defaultdict(float)
    for e in entries:
        month_totals[e[0]] += e[3]
        cat_totals[e[1]] += e[3]
    
    print("\n=== By Month ===")
    for m in sorted(month_totals.keys()):
        print(f"  {m}: R{month_totals[m]:,.2f}")
    
    print("\n=== By Category ===")
    for cat, total in sorted(cat_totals.items(), key=lambda x: x[1], reverse=True):
        print(f"  {cat}: R{total:,.2f}")
    
    # Import to database
    conn = sqlite3.connect(DB_PATH)
    reset_cc_imported_rows(conn)
    insert_cc_entries(conn, entries)
    rebuild_cc_snapshots(conn)
    conn.close()
    
    print(f"\n✅ Successfully imported {len(entries)} credit card transactions")
    print(f"   Source: fnb_credit_card")
    print(f"   Excluded: {excluded_count} transactions (business/transfers)")


if __name__ == "__main__":
    main()