---
name: bulk-status
description: >
  Match a vendor/venue's invoice-and-RFP overview against Business Central's
  Vendor Ledger Entries export and produce a status-update Excel per the Bulk
  Status SOP. Use whenever Rogier gets a "bulk status" request — an email or
  Excel from a venue (Van der Valk, BCN, Park Plaza, Hilton, or ANY other
  venue) asking for a status update on several invoices/RFPs at once — and
  wants those checked against Business Central. Trigger on "bulk status",
  "status update for these invoices", "check these RFPs against BC/Business
  Central", "compare against Vendor Ledger Entries", or when a Vendor Ledger
  Entries export is uploaded alongside a file with invoice numbers and RFP
  references. NOT limited to the SOP's named venues — any Excel with an
  invoice/document number plus an RFP reference (own column, or trailing
  digits in a description field) qualifies. Outputs a new Excel with match
  status and a plain-language status per row; never sends anything.
---

# Bulk Status

Weekly-ish task: a venue emails asking "what's the status of these invoices/RFPs",
usually as an Excel with one row per invoice. Instead of manually building
`INDEX`/`MATCH` formulas against the Business Central export every time (the
original SOP's manual method), run the matcher script and hand back a finished,
color-coded status workbook.

## What the SOP formulas actually do

The SOP's Dutch/English `INDEX`/`MATCH` formulas boil down to this, cross-checked
in both directions so a wrong invoice number on the venue's side doesn't slip
through silently:

1. **Direct match**: venue's invoice number → Business Central `External Document No.`
   Pull back BC's `RFP No.`, `Remaining Amount`, `Due Date`, `Open`.
2. **RFP match**: RFP number → BC `RFP No.` → pull back BC's `External Document No.`
   The RFP number is usually the trailing digit run of a description/reference
   field (e.g. `...MEETINGSELECT B.V. RFP ID: 759257` → `759257`), unless the
   file has its own dedicated RFP column.
3. **Cross-check**: do the two paths agree? If BC's RFP-via-invoice-number
   doesn't equal the RFP-via-description, or vice versa, that's a **discrepancy**,
   not a silent match — flag it rather than trusting one lookup.
4. **Payment status**, only once matched:
   - BC `Open` = 0/false → **Paid**
   - BC `Open` = 1/true and `Due Date` in the past → **Overdue – payment pending**
   - BC `Open` = 1/true and `Due Date` not yet reached → **Open – not yet due**
   - No match on either path at all → **Not found in our records**

`scripts/match_status.py` implements exactly this, with header-name based column
detection (not hardcoded cell references), so it works on any venue's file as
long as it has an invoice-number-like column and either an RFP column or a
description field ending in the RFP number.

## Workflow

1. **Identify the two inputs.** Business Central's Vendor Ledger Entries export
   (has columns like `External Document No.`, `RFP No.`, `Remaining Amount`,
   `Due Date`, `Open`) and the venue's overview (has an invoice number column
   and either an RFP column or a description with the RFP number trailing it).
   If either file is missing or ambiguous, ask before guessing — don't run the
   matcher against the wrong pair of files.

2. **Run the matcher:**
   ```bash
   python scripts/match_status.py <bc_export.xlsx> <vendor_file.xlsx> <output_dir>
   ```
   Optionally pass `--today YYYY-MM-DD` when testing against historical data so
   overdue/open calculations use the right reference date; omit it in real use
   (defaults to today).

   The third argument is a **directory**, not a filename — see naming
   convention below.

3. **Sanity-check before sending anything back:**
   - Read the printed matched / discrepancy / not-found counts.
   - Spot check a few rows of each category with `openpyxl` — a "Discrepancy"
     row deserves a second look before it goes in front of a venue, since it
     usually means either their invoice number or their RFP number is wrong.
   - If discrepancies or not-found rows are a large fraction of the file,
     say so plainly rather than shipping the workbook quietly — that's usually
     a column-detection problem (wrong file, unexpected header names) rather
     than genuinely unmatched invoices.

4. **Deliver the output workbook** — it has every original column plus:
   `RFP No. (extracted)`, `BC: External Doc. No. (via RFP)`, `Match check 1`,
   `BC: RFP No. (via invoice no.)`, `Match check 2`, `Match status`,
   `BC: Remaining Amount`, `BC: Due Date`, `Status update` (color-filled:
   green = Paid, amber = Open/not yet due, red = Overdue, grey = Not found).

## Naming convention (always applies, no exceptions)

Output filename is always:

```
Status Overview - <Vendor Name> - <Vendor No>.xlsx
```

e.g. `Status Overview - Van der Valk Hotel Lelystad A6 - V02654.xlsx`. This is
not something you choose or ask about — the script derives it automatically
from Business Central's `Vendor No.` / `Vendor Name` on the matched rows, and
always writes into the output directory you pass, never a path you name
yourself. If you're presenting the file to Rogier rather than just running
the script, keep this exact filename — don't rename it to something more
"readable" on the way out.

**If the matched rows span more than one vendor**, there's no single
name/number to put in the filename, so it becomes:

```
Status Overview - Multiple Venues.xlsx
```

That's usually worth a second look before shipping — it typically means the
wrong two files were paired up (a vendor file matched against a BC export
covering several vendors at once) rather than a genuinely mixed request.

## The status wording is a house convention, not gospel

The four `Status update` phrases (`Paid`, `Open – not yet due`,
`Overdue – payment pending`, `Not found in our records`) are the current
default — confirm with Rogier if a specific venue reply needs different
wording, and update the `STATUS_*` constants in `match_status.py` rather than
hand-editing output files row by row.

## Guardrails

- Never invent a match. If neither the invoice number nor the RFP number is
  found in Business Central, the row is "Not found" — don't guess based on
  vendor name or amount alone.
- Don't silently resolve a discrepancy by picking whichever lookup "looks more
  right." Surface it — that's the entire point of doing both directions.
- This skill only produces the Excel. It never drafts or sends the reply email
  to the venue — that's a separate step (use `me-style` or `rogier-ms-voice`
  if Rogier wants that drafted afterwards).
- If the vendor file's invoice-number or description column can't be detected
  by header name, don't force a guess — report which headers were found and
  ask which column to use.
