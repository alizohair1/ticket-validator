# TicketLens — WhatsApp Ticket Validator

Validates WhatsApp feedback tickets against a spooler file.

## How it works

Upload two files:
1. **Spooler file** — contains the orders sent (columns: DEST, CREATED AT)
2. **Tickets file** — contains customer feedback (columns: Ticket name / Customer CLI, Created At)

### Validation rules

| Rule | Result |
|------|--------|
| Phone exists in spooler AND date matches | ✅ Valid |
| Phone not found in spooler | ❌ Invalid — "Phone not in spooler" |
| Phone found but date doesn't match | ❌ Invalid — "Date mismatch (late feedback)" |
| Duplicate ticket (same phone + same date) | ❌ Invalid — "Duplicate ticket" |

### Output
The result keeps the original ticket structure and adds two columns:
- **Status** — Valid / Invalid
- **Reason** — reason for invalid (empty for valid)

Download as XLSX or CSV.

## Getting started

```bash
npm install
npm run dev
```

## Deploy to Vercel

1. Push this repo to GitHub
2. Go to [vercel.com](https://vercel.com) → New Project → Import from GitHub
3. Vercel auto-detects Vite — click **Deploy**
4. Done! Your app is live.

No backend or environment variables needed — runs 100% in the browser.
