# House Format — Canonical Examples

These are the two category blocks already accepted. Match their tone, density, and
structure. Note: English prompt body, bilingual cues, a discriminator line, the
fixed closing sentence, then a Dutch "distilled" note.

## Example — Payment Status

**✦ Je prompt:**

> Classify the email as **Payment Status** if it contains any of the following signals:
>
> - The sender explicitly asks when Meetingselect will pay an outstanding invoice
> - The email is a payment reminder (herinnering, betalingsherinnering) or dunning notice (aanmaning)
> - The email references an overdue amount, due date, or threatens consequences for non-payment (service restriction, credit bureau, legal action)
> - The email follows up on a previously sent invoice that has not yet been paid
>
> These signals may appear in the subject line, body, or attachment content. Language is typically Dutch or English.

---

**Wat ik eruit heb gedistilleerd:** Alleen de detectielogica — geen JSON, geen
output-format, geen voorbeelden. Plak dit als conditieblok in de bredere prompt.

## Example — Multiple Statusses Sheet

**✦ Je prompt:**

> Classify the email as **Multiple Statusses Sheet** if it contains any of the following signals:
>
> - The email concerns **multiple** open or overdue items at once (two or more invoice numbers, RV numbers, RFP IDs, or bookings) that call for a single consolidated status overview — one status per line
> - It includes, references, or attaches an overview, statement, or sheet (Excel/.xlsx, "Statement", "overzicht", "specificatie"), an inline column-table of invoices/amounts, or a screenshot of a previously shared sheet
> - The sender asks to reconcile records, verify details, or provide a status per item ("reconcile our records", "verify your details", "graag per factuur een status", "kunt u dit afstemmen")
> - It is a dunning/aanmaning covering several overdue invoices with a combined total ("totaalbedrag aan vervallen facturen € …"), rather than a single invoice
>
> Plurality is decisive: a single-invoice inquiry, reminder, or dunning is **Payment Status**, not this category. When both seem to fit, choose **Multiple Statusses Sheet** if a consolidated status sheet is the natural reply.
>
> These signals may appear in the subject line, body, or attachment content. Language is typically Dutch or English.

---

**Wat ik eruit heb gedistilleerd:** Alleen detectielogica als conditieblok. De
plurality-regel is bewust toegevoegd omdat dat het enige echte breekpunt is met
Payment Status; zonder die regel haalt het model die twee door elkaar.
