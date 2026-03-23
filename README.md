# OpenMerge

> It is fully open source, runs entirely inside your Google account, and replicates everything these paid tools charge you for — at zero cost.

**No subscription. No data leaving your account. No vendor lock-in. Just Google Sheets, Gmail, and a script that works.**

[![License: Apache 2.0](https://img.shields.io/badge/License-Apache%202.0-blue.svg)](https://opensource.org/licenses/Apache-2.0)
[![Platform: Google Apps Script](https://img.shields.io/badge/Platform-Google%20Apps%20Script-green.svg)](https://script.google.com)
[![Free Forever](https://img.shields.io/badge/Price-Free%20Forever-brightgreen.svg)]()

---

## Why OpenMerge over YAMM, Mailmeteor, or GMass?

| Feature | OpenMerge | YAMM | Mailmeteor | GMass |
|---|---|---|---|---|
| Free to use | Always | 50 emails/day cap | 150 emails/day cap | 25 emails/day cap |
| Open source | Yes | No | No | No |
| Your data stays in your account | Always | Third-party servers | Third-party servers | Third-party servers |
| Open tracking | Yes | Paid only | Paid only | Paid only |
| Bounce detection | Yes | Paid only | Paid only | Paid only |
| Campaign summary dashboard | Yes | Paid only | Paid only | Paid only |
| Quota management | Built-in | No | No | No |
| Multi-day sending | Auto-resume | Manual | Manual | Manual |
| Customisable | Full source | No | No | No |

OpenMerge runs entirely inside your Google account using Apps Script. Your recipient data, email content, and tracking results never touch a third-party server.

---

## Features

| Feature | Description |
|---|---|
| **Mail Merge** | Send personalised emails using Gmail drafts as templates with `{{ColumnName}}` placeholders |
| **Open Tracking** | Pixel-based tracking logs open count and last opened timestamp per recipient |
| **Bounce Detection** | Scans Gmail for delivery failure notifications across all major mail servers and marks bounced rows |
| **Quota Management** | Displays remaining daily quota before every send run using Gmail's live quota API |
| **Multi-day Sending** | Automatically resumes from where it left off the next day when quota is reached |
| **Campaign Dashboard** | Live summary with metric cards, progress bars, and a per-recipient status table |
| **Auto Column Setup** | Missing sheet columns are created and formatted automatically on open |
| **Debug Tool** | Diagnoses placeholder mismatches between your Gmail draft and sheet headers |

---

## Table of Contents

1. [Prerequisites](#prerequisites)
2. [Sheet Setup](#sheet-setup)
3. [Installation](#installation)
4. [Web App Deployment](#web-app-deployment-open-tracking)
5. [Usage](#usage)
6. [Column Reference](#column-reference)
7. [Menu Reference](#menu-reference)
8. [Email Template Syntax](#email-template-syntax)
9. [Sending Limits](#sending-limits)
10. [Tracking Limitations](#tracking-limitations)
11. [Troubleshooting](#troubleshooting)

---

## Prerequisites

- A Google account (free Gmail `@gmail.com` or Google Workspace)
- A Google Sheet with at least an **Email** column containing recipient addresses
- A Gmail draft to use as the email template

---

## Sheet Setup

Your sheet must have the following column headers in **row 1**. They will be created and formatted automatically if missing when the sheet is opened after installation.

| Column Header | Purpose |
|---|---|
| `Email` | Recipient email addresses |
| `Merge Status` | Tracks send status per row (blank = not yet sent) |
| `Opens` | Running count of times each email was opened |
| `Last Opened` | Date and time of the most recent open event |
| `Bounced` | `TRUE` / `FALSE` after a bounce check is run |

Any additional columns such as `First Name`, `Company`, or `Description` can be used as merge fields in your template.

---

## Installation

1. Open your Google Sheet.
2. Click **Extensions → Apps Script**.
3. Delete any existing code in the editor.
4. Paste the full `OpenMerge.gs` script into the editor.
5. Confirm the constants at the top of the script match your column headers:
```javascript
const RECIPIENT_COL  = "Email";
const EMAIL_SENT_COL = "Merge Status";
const OPEN_COUNT_COL = "Opens";
const LAST_OPEN_COL  = "Last Opened";
const BOUNCE_COL     = "Bounced";
```

6. Save the project (`Ctrl + S` or `Cmd + S`).
7. Close and reopen your Google Sheet — a **Mail Merge** menu will appear in the toolbar.

---

## Web App Deployment (Open Tracking)

Open tracking requires the script to be deployed as a public Web App. The tracking pixel embedded in each email calls this URL when a recipient opens the message.

1. In the Apps Script editor, click **Deploy → New Deployment**.
2. Click the gear icon next to **Type** and select **Web App**.
3. Set **Execute as** to `Me`.
4. Set **Who has access** to `Anyone`.
5. Click **Deploy** and copy the generated URL.
6. Paste the URL into the script constant:
```javascript
const TRACKING_BASE_URL = "https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec";
```

7. Save the script.

> **Important:** If you modify the script after deployment, you must create a **New Deployment** rather than updating the existing one for changes to take effect. Copy the new URL and update `TRACKING_BASE_URL`.

---

## Usage

### Sending Emails

1. Ensure your sheet has recipient data and your Gmail draft is ready.
2. Click **Mail Merge → Send Emails**.
3. A quota check dialog will appear showing how many emails you can send today — click **OK** to proceed.
4. Enter the **exact subject line** of your Gmail draft when prompted.
5. Emails are sent to all rows where `Merge Status` is blank. Rows already stamped are skipped automatically.

### Resuming the Next Day

When the daily quota is reached mid-run, affected rows are left blank (not stamped with an error) so the script picks them up automatically on the next run. The post-send summary dialog shows exactly how many remain pending.

### Checking Bounces

1. Click **Mail Merge → Check Bounces**.
2. The script searches your Gmail inbox for delivery failure notifications from all major mail servers (Google, Outlook, Yahoo, corporate servers).
3. The `Bounced` column is updated — `TRUE` for confirmed bounces, `FALSE` for successfully delivered rows.
4. Colour formatting is applied: pale red for bounced, pale green for delivered.

### Viewing the Campaign Summary

1. Click **Mail Merge → Campaign Summary**.
2. A prompt asks whether to run a bounce check first — recommended for accurate results.
3. The dashboard displays:
   - Metric cards: **Delivered**, **Opened**, **Bounced**, **Pending**
   - Progress bars showing percentage breakdowns
   - A per-recipient table with colour-coded status badges

### Checking Your Quota

Click **Mail Merge → Check Quota** at any time to see how many emails you have remaining today without triggering a send.

---

## Column Reference

### Merge Status Values

| Value | Meaning |
|---|---|
| *(blank)* | Not yet sent — will be included in the next send run |
| `Sent: DD/MM/YYYY` | Successfully sent on the given date |
| `Error: [message]` | A real error occurred (e.g. invalid address) — will not be retried automatically |

### Bounced Values

| Value | Meaning | Cell Colour |
|---|---|---|
| *(blank)* | Bounce check has not been run for this row | No fill |
| `FALSE` | Sent and no bounce detected | Pale green |
| `TRUE` | Delivery failure confirmed | Pale red |

---

## Menu Reference

| Menu Item | Function |
|---|---|
| **Send Emails** | Sends to all unsent rows, respecting the daily quota limit |
| **Check Bounces** | Scans inbox for delivery failures and updates the Bounced column |
| **Check Quota** | Shows remaining daily email quota without triggering a send |
| **Debug Placeholders** | Diagnoses mismatches between draft placeholders and sheet column headers |
| **Campaign Summary** | Opens the live campaign dashboard with progress bars and per-row status |

---

## Email Template Syntax

Merge fields are written as `{{ColumnName}}` in your Gmail draft and are replaced with the corresponding cell value for each recipient at send time.

**Example draft body:**
```
Hi {{First Name}},

I wanted to reach out to you at {{Company}} regarding...

Best regards
```

### Rules

- The placeholder must **exactly** match the sheet column header — including capitalisation and spacing.
- Write placeholders in **plain text**. Gmail's rich text editor can inject HTML `<span>` tags inside `{{}}` when text is pasted or styled, which will silently break the merge.
- If placeholders are not resolving, run **Debug Placeholders** from the menu. It will show you the exact mismatch. Then retype the placeholder from scratch in your Gmail draft with all formatting removed (`Format → Remove formatting` or `Ctrl+\`).

---

## Sending Limits

OpenMerge works within Gmail's own sending limits — no artificial caps imposed on top.

| Account Type | Daily Limit |
|---|---|
| Free Gmail (`@gmail.com`) | 100 emails/day |
| Google Workspace | 1,500 emails/day |

Limits reset at **midnight Pacific Time**. The script reads your live remaining quota via `MailApp.getRemainingDailyQuota()` before every send run so you always know exactly where you stand.

For lists larger than your daily limit, run **Send Emails** again the following day. Completed rows are stamped and skipped — OpenMerge picks up exactly where it left off.

---

## Tracking Limitations

### Open Tracking

Open tracking works by embedding a 1×1 invisible pixel in the HTML email body. When the recipient opens the email, their email client loads the pixel from the Web App URL, which logs the event in the sheet.

This is the same mechanism used by YAMM, Mailmeteor, and GMass. It has the following known limitations:

- **Image blocking** — email clients that block remote images will not trigger the pixel.
- **Apple Mail Privacy Protection** — from iOS 15 / macOS Monterey onwards, Apple Mail pre-fetches all remote images on delivery, which may log an open before the recipient actually reads the message.
- **Plain-text clients** — recipients reading a plain-text version of the email will not trigger the pixel.

Open counts should be treated as a directional signal rather than a precise figure. This is true of every mail merge tracking tool on the market.

### Bounce Detection

Bounce detection searches your Gmail inbox for delivery status notifications. It will not detect bounces if:

- The bounce notification email has been deleted or archived from your inbox before the check is run.
- The receiving mail server silently drops the message without sending a failure notification (silent drop / soft fail).
- Your inbox has more than 50 recent bounce notification threads (the search is capped at 50 threads per query per search pattern).

---

## Troubleshooting

### Placeholders are not replacing in sent emails

Run **Mail Merge → Debug Placeholders** and enter your draft subject line. The tool reports whether each `{{placeholder}}` in the draft matches a column header exactly.

The most common cause is Gmail injecting HTML `<span>` tags inside `{{}}` when text is pasted or formatted. Fix this by selecting all text in the draft, removing all formatting (`Format → Remove formatting` or `Ctrl+\`), and retyping the placeholder from scratch in plain text.

### Bounced column is showing incorrect results

Ensure:
1. The row has a `Sent:` value in `Merge Status` — bounce checking only processes rows that were actually sent.
2. You re-run **Check Bounces** and then open **Campaign Summary** fresh after the check completes.

### Emails are not sending — quota error

Your daily sending limit has been reached. Affected rows are left blank and will be picked up automatically on the next run. Check your remaining quota at any time with **Mail Merge → Check Quota**.

### The Web App URL is returning an error after editing the script

After modifying the script, create a **New Deployment** from **Deploy → New Deployment**. Copy the new URL and update `TRACKING_BASE_URL` in the script constants. Updating an existing deployment does not always propagate changes to the live Web App endpoint.

### The Mail Merge menu does not appear after installation

The `onOpen` trigger creates the menu when the sheet is opened. If it is missing, go to **Extensions → Apps Script**, select `onOpen` from the function dropdown at the top of the editor, and click **Run** once to initialise it manually. Then reopen the sheet.

---

## License

Licensed under the [Apache License 2.0](https://opensource.org/licenses/Apache-2.0).

You are free to use, modify, and distribute this script for personal or commercial purposes. Attribution is appreciated but not required.

---

## Contributing

Pull requests are welcome. If you have encountered a bounce notification format not currently detected, a new mail server pattern, or any other improvement — open an issue or submit a PR.

---

*Built out of frustration with paywalled tools. Free forever.*
