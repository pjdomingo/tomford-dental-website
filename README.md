# TomFord Dental — Official Website

Clinic website for **TomFord Dental** (Pasig, Philippines).
*Where every tooth matters.*

Live site: _pending custom domain_

---

## Stack

- **Frontend** — single-page `index.html` (Tailwind via CDN, vanilla JS, Cormorant Garamond + Inter)
- **Hosting** — Cloudflare Pages (or Workers static assets)
- **Booking backend** — Google Apps Script Web App backed by a Google Sheet + Google Calendar + Gmail
- **DNS / CDN** — Cloudflare

## Repo layout

```
.
├── index.html          # the website (all markup, styles, JS inlined)
├── _headers            # Cloudflare security headers + CSP
├── .assetsignore       # files excluded from the Worker/Pages deploy
├── apps-script/
│   ├── Code.gs         # Google Apps Script backend — bookings, slots, email
│   └── SETUP.md        # deployment & environment setup
└── README.md
```

## Deploying the website

1. Push to `main` — Cloudflare Pages auto-deploys the root of the repo.
2. `_headers` is applied by Cloudflare automatically.
3. `apps-script/`, `README.md`, and `.github/` are excluded from the deploy via `.assetsignore`.

## Deploying the booking backend

See [`apps-script/SETUP.md`](apps-script/SETUP.md) for the full walkthrough. TL;DR:

1. Open [script.google.com](https://script.google.com), create a new project, paste `Code.gs`.
2. Set Script Properties: `SHEET_ID`, `CALENDAR_ID`, `CLINIC_EMAIL`, `TIMEZONE=Asia/Manila`.
3. Deploy as **Web App** → **Execute as: Me** → **Anyone**.
4. Copy the deployed URL and paste into `index.html` as the `scriptURL` constant.
5. Re-deploy the website.

> Important: the web app **must** be deployed as "Execute as: Me" for the script to reach the clinic's Sheet, Calendar, and Gmail. "User accessing" will break everything.

## Editing content

All copy, services, pricing, and imagery live in `index.html`. Edit, commit, push — it's live in ~1 minute.

The booking modal, slot logic, and email templates are in the same file for now — easy enough to grep, split later if it grows.

## Support

Paolo Domingo — domingopauljohn@gmail.com
