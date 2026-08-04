# AP Inbox Category Map

The classifier sorts every email to `accountspayable@meetingselect.com` into
exactly one of five categories. For each, the **defining trait** is what to detect;
the **breakpoint** is the adjacent category it is most often confused with. Mirrors
the AR inbox setup; thread linking uses Conversation ID (primary) + Message ID.

## Payment Status
**Defining trait:** A supplier is asking about / chasing payment of an invoice
Meetingselect owes — a single outstanding or overdue invoice.
**Signals (already drafted):** explicit "when will you pay" inquiries; payment
reminders (herinnering, betalingsherinnering) or dunning notices (aanmaning);
references to an overdue amount, due date, or threatened consequences (service
restriction, credit bureau, legal action); follow-ups on an unpaid invoice.
Also covers: reminder escalation/sequence numbers, automated no-reply accounting
platforms, internally forwarded (FW:) reminders, and a payment request buried in
an ongoing RFP/reservation thread without "herinnering" in the subject.
**Breakpoint:** vs **Multiple Statusses Sheet** — ONE invoice → Payment Status;
MULTIPLE items needing a consolidated per-line status → Multiple Statusses Sheet.
**Downstream:** extract identifiers (RFP ID, Invoice Number), capture sender,
persist for validation. Execution scenarios tracked in a separate ticket.

## Corrections
**Defining trait:** The sender requests an adjustment or correction to an invoice
or a specific detail on it. Classify and store only — no automated action.
**Breakpoint:** vs Payment Status — a correction is about *fixing* the invoice,
not chasing its payment.

## Reminders
**Defining trait:** A response to a Meetingselect request asking a supplier to
upload an invoice into the Meetingselect platform. Classify and store only.
**Breakpoint:** vs Payment Status — here Meetingselect initiated the chase
(supplier replying to *our* request); in Payment Status the supplier is chasing us.

## Multiple Statusses Sheet
**Defining trait:** The email concerns MULTIPLE open/overdue items at once
(several invoice numbers, RV numbers, RFP IDs, or bookings) that call for a single
consolidated status overview — one status per line. (Formerly "Van der Valk Sheet";
renamed because it's about giving multiple statuses in a sheet, not one venue.)
**Signals:** multiple itemised items in one message; an attached/inline/referenced
overview, statement, or sheet (Excel/.xlsx, "Statement", "overzicht",
"specificatie", a column-table, or a screenshot of a prior sheet); a request to
reconcile/verify/status the whole batch; an account-level dunning across several
overdue invoices with a combined total.
**Breakpoint:** plurality is decisive vs Payment Status. When both fit, choose
Multiple Statusses Sheet if a consolidated status sheet is the natural reply.
**Reply logic (separate component — do NOT put in detection):** check for Excel
attachment; if missing, evaluate whether one is required (an inline table or a
screenshot referencing a prior sheet counts as valid); if required but absent,
reply with the designated template; if present, reply with a formatted Excel
(Status in column K, Toelichting in column L; codes `*` = Dispuut, `T` = Toezegging).

## Unidentifiable
**Defining trait:** Catch-all — cannot be confidently matched to any category
above. Use only after the other four have been ruled out.
**Breakpoint:** prefer a real category whenever one clearly fits; Unidentifiable
is the last resort, not a tie-breaker.
