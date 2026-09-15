"""Standing post-import verification gate for the personal budget DB.

Run after EVERY import:  python3 check_budget.py
Exit code != 0 means the import broke something.

Counts are FLOORS, not equality checks, so importing new statement periods
(August onward) does not false-fail. Raise the floors as scope grows.
Closed-month expectations (like March 2026) stay exact forever.
"""
import sqlite3
import sys

DB = '/home/molty/Projects/budget-dashboard/budget.db'
conn = sqlite3.connect(DB)
failures = []


def check(label, ok, detail=''):
    print(('  PASS ' if ok else '  FAIL ') + label + (f' | {detail}' if detail else ''))
    if not ok:
        failures.append(label)


# --- Checkpoints match live entries ---
live_rows = conn.execute("""
    SELECT month,
           SUM(CASE WHEN entry_type='income' THEN amount ELSE 0 END),
           SUM(CASE WHEN entry_type='expense' THEN amount ELSE 0 END),
           SUM(CASE WHEN entry_type='savings' THEN amount ELSE 0 END)
    FROM budget_entries GROUP BY month ORDER BY month""").fetchall()
live = {m: (i or 0, e or 0, s or 0) for m, i, e, s in live_rows}

snap = {r[0]: (r[1], r[2], r[3]) for r in conn.execute(
    'SELECT month, total_income, total_expenses, total_savings FROM monthly_snapshots').fetchall()}
bad = [m for m, t in live.items()
       if m not in snap or any(abs(a - b) > 0.005 for a, b in zip(snap[m], t))]
check('monthly_snapshots match live budget_entries', not bad,
      f'months={len(live)} bad={bad[:5]}')

march = snap.get('2026-03', (0, 0, 0))
check('March 2026 checkpoint expenses == 58,694.96 (rebate -31.60 + Sanlamgap credit -291.00)',
      abs(march[1] - 58694.96) < 0.005, f'R{march[1]:,.2f}')

april = snap.get('2026-04', (0, 0, 0))
check('April 2026 checkpoint expenses == 43,656.00 (Uys Burger stock deposit removed)',
      abs(april[1] - 43656.00) < 0.005, f'R{april[1]:,.2f}')

may = snap.get('2026-05', (0, 0, 0))
check('May 2026 checkpoint expenses == 72,234.57 (closed month)',
      abs(may[1] - 72234.57) < 0.005, f'R{may[1]:,.2f}')

june = snap.get('2026-06', (0, 0, 0))
check('June 2026 checkpoint expenses == 70,877.38 (rebate -5.72)',
      abs(june[1] - 70877.38) < 0.005, f'R{june[1]:,.2f}')

july = snap.get('2026-07', (0, 0, 0))
check('July 2026 checkpoint expenses == 34,008.58 (rebate -13.46)',
      abs(july[1] - 34008.58) < 0.005, f'R{july[1]:,.2f}')

dupes = conn.execute('SELECT month, COUNT(*) c FROM monthly_snapshots GROUP BY month HAVING c>1').fetchall()
check('no duplicate snapshot months', not dupes, str(dupes))

# --- Gerhard's classification rulings hold ---
n, tot, cats = conn.execute("""
    SELECT COUNT(*), ROUND(SUM(amount),2), GROUP_CONCAT(DISTINCT category)
    FROM budget_entries WHERE source='fnb_credit_card'
    AND LOWER(subcategory) LIKE '%xtra savings%'""").fetchone()
check('Xtra Savings included as subscription (>=6 rows, >=R594)',
      n >= 6 and tot >= 594 and cats == 'subscription', f'{n} rows R{tot} {cats}')

sl = conn.execute("""SELECT COUNT(*) FROM budget_entries
    WHERE LOWER(subcategory) LIKE '%slow lounge%'
       OR LOWER(subcategory) LIKE '%slow entry fee reversal%'""").fetchone()[0]
check('fully-reversed fees stay netted out (Slow Lounge)', sl == 0, f'{sl} rows')

leaks = []
for pat in ['%cum tygervalley%', '%payfast%creative%', '%amazon retail%', '%boyztoyz%',
            '%pudo%', '%maseeg damon%', '%lochner%', '%wa burger kaalplek%', '%aloha holdings%',
            '%mohamed badrudin%', '%janse van rensbur%', "%huawei gt5%", "'%p4p%'",
            '%uys burger%', '%dr interest rebate%', '%rtc credit%']:
    pat = pat.strip("'")
    c = conn.execute('SELECT COUNT(*) FROM budget_entries WHERE LOWER(subcategory) LIKE ?',
                     (pat,)).fetchone()[0]
    if c:
        leaks.append((pat, c))
check('business stock/courier merchants + inflows still excluded', not leaks, str(leaks))

generic_notes = conn.execute("""SELECT COUNT(*) FROM budget_entries
    WHERE source IN ('fnb_statement','fnb_credit_card')
    AND notes LIKE 'Imported from%'""").fetchone()[0]
check('no boilerplate import notes remain', generic_notes == 0, f'{generic_notes} rows')

distinct_notes = conn.execute("""SELECT COUNT(DISTINCT notes) FROM budget_entries
    WHERE source IN ('fnb_statement','fnb_credit_card')""").fetchone()[0]
check('notes explain items (>=25 distinct)', distinct_notes >= 25, f'{distinct_notes} distinct')

# --- Structural invariant: no non-salary INFLOW may be booked as an expense ---
import json as _json
import os as _os
_feed_path = '/home/molty/transactions.json'
if not _os.path.exists(_feed_path):
    check('inflow-leak feed present (transactions.json)', False,
          'missing - re-link the finance-archive symlink so this invariant can run')
else:
    _raw = _json.load(open(_feed_path))
    _inflows = {(str(t['description']).strip().lower(), round(abs(float(t['amount'])), 2))
                for t in _raw if float(t['amount']) > 0
                and 'salary' not in str(t.get('description', '')).lower()}
    _booked = set(conn.execute("""SELECT LOWER(subcategory), ROUND(amount, 2)
        FROM budget_entries WHERE source='fnb_statement' AND entry_type='expense'""").fetchall())
    _leaked = _booked & _inflows
    check('no business inflow booked as cheque expense',
          not _leaked, f'{len(_leaked)} e.g. {list(_leaked)[:2]}')

# --- Completeness floors ---
cnts = dict(conn.execute("""SELECT source, COUNT(*) FROM budget_entries
    WHERE source IN ('fnb_statement','fnb_credit_card') GROUP BY source""").fetchall())
check('source floors met (>=152 cheque / >=175 credit card)',
      cnts.get('fnb_statement', 0) >= 152 and cnts.get('fnb_credit_card', 0) >= 175, str(cnts))

non_sal = conn.execute("""SELECT COUNT(*) FROM budget_entries
    WHERE entry_type='income' AND source IN ('fnb_statement','fnb_credit_card')
    AND subcategory NOT LIKE '%Salary%'""").fetchone()[0]
check('income = salary deposits only', non_sal == 0, f'{non_sal} non-salary income rows')

conn.close()
print()
if failures:
    print(f'RESULT: {len(failures)} FAILED -> {failures}')
    sys.exit(1)
print('RESULT: ALL CHECKS PASSED')
