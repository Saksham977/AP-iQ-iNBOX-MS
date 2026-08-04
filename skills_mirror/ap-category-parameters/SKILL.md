---
name: ap-category-parameters
description: >
  Write or update the detection parameters (the "Classify the email as X if..."
  condition block) for a category of the Meetingselect AP inbox classifier
  (accountspayable@meetingselect.com). Use this skill whenever Rogier uploads
  real supplier emails (.eml or .msg) and wants the detection logic for an AP
  category — Payment Status, Corrections, Reminders, Multiple Statusses Sheet,
  or Unidentifiable — or asks to "write parameters", "update the prompt for
  category X", "how do we identify these as category Y", or "make detection
  parameters from these emails". Trigger this even when the word "parameters"
  isn't used: if supplier emails are provided alongside a category name and the
  goal is classification logic, this skill applies. Produces parameters in the
  fixed house format, matching the style used for Payment Status and Multiple
  Statusses Sheet. Do NOT use for the full classification prompt (JSON, examples),
  the reply/execution step, or the weekly reports — detection logic only.
---

# AP Inbox Category — Detection Parameters

This skill produces one thing: a clean, paste-ready **detection block** that tells
the AP inbox classifier (GPT-4.1 mini) how to recognise a single category. It is
deliberately narrow — only the logic that distinguishes one category from the
others, in a consistent house format that Rogier assembles into the broader
classification prompt later.

## Why the format is what it is

Rogier builds the classifier one category at a time and strips each component down
to a reusable block. The parameters must be:
- **English prose** (the prompt language) with **Dutch *and* English signal cues**,
  because suppliers write in both languages.
- **Detection only** — no JSON, no output schema, no few-shot examples. Those live
  elsewhere in the prompt.
- **Discriminating** — the single hardest job is preventing overlap with the
  *adjacent* category. Always name the breakpoint explicitly.

## Workflow

### 1. Read the uploaded emails properly

`.eml`/`.msg` files are not in context — parse them, don't `cat` them. Extract
subject, From, To, body, and attachment filenames/types. Attachments matter:
the presence of an Excel/statement/PDF set or an inline table is itself a signal.

```python
import email
from email import policy
with open(path, 'rb') as f:
    msg = email.message_from_binary_file(f, policy=policy.default)
# headers: Subject, From, To, Date, Message-ID, In-Reply-To, References
# body: msg.get_body(preferencelist=('plain','html')).get_content()
# attachments: walk parts, collect get_filename() + get_content_type()
```
For `.msg`, use `extract-msg` (`pip install extract-msg --break-system-packages`)
or convert first. If a body is HTML-only, strip tags before reading.

### 2. Identify the core discriminator

Before listing signals, settle the ONE trait that makes this category itself and
not its neighbour. Read `references/category-map.md` for each category's defining
trait and the most-confusable pairing. The discriminator goes at the top of the
block (or as the closing rule) — it is the part that actually prevents misfires.

### 3. Extract bilingual signal cues from the real examples

Pull the concrete phrases that appear in the uploaded emails — Dutch and English —
and generalise them into 3–5 signal patterns. Each pattern should be a behaviour
("the sender asks to reconcile multiple invoices"), with concrete cue words in
parentheses drawn from the actual emails. Don't invent cues that aren't grounded
in a real example or a clear variant of one.

### 4. Write the block in the house format

ALWAYS use this exact structure. See `references/house-format.md` for the two
canonical outputs (Payment Status, Multiple Statusses Sheet) to match tone and
density.

```markdown
**✦ Je prompt:**

> Classify the email as **<Category>** if it contains any of the following signals:
>
> - <signal pattern 1, with (Dutch cue, English cue)>
> - <signal pattern 2 …>
> - <signal pattern 3 …>
> - <signal pattern 4 …>
>
> <One discriminator line naming the breakpoint vs the adjacent category.>
>
> These signals may appear in the subject line, body, or attachment content. Language is typically Dutch or English.

---

**Wat ik eruit heb gedistilleerd:** <Dutch note: what was stripped (JSON/format),
which uploaded examples the signals cover, and any rule added on purpose and why.>
```

### 5. Updating an existing category

If parameters for this category already exist (search past chats in this project
if unsure), don't rewrite from scratch. Compare the new emails against the current
signals, keep what holds, and add only the genuinely new patterns — then say in
the distilled note which patterns are net-new and which uploaded email prompted
each. This mirrors how the Payment Status block grew over several email batches.

## Guardrails

- Never widen a category so far it swallows its neighbour. When two categories
  both seem to fit, the discriminator decides — state which wins and why.
- Keep it to 3–5 signals. More than that usually means two categories are being
  merged, or a signal is really an example, not a pattern.
- Don't add the reply/execution logic (e.g. which Excel columns to send back).
  That is a separate component tracked elsewhere — stop at detection.
- Ground every cue in the uploaded material. If a category needs a signal no
  example demonstrates, flag it as an assumption rather than presenting it as
  observed.
