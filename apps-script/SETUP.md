# TomFord Dental — Apps Script Setup

## 1. Open the Google Sheet
Open the existing Google Sheet connected to the booking form.

## 2. Open Apps Script
Go to **Extensions → Apps Script**.  
Replace ALL existing code with the contents of `Code.gs`.

## 3. Create these Sheet tabs

### Tab: `Services`
Column A — one service per row (no header row needed):
```
Consultation
TMJ Consultation
Periapical X-ray
Teeth Whitening
Jacket Crown
Fixed Bridge
Veneers
Dentures
Root Canal Treatment
Oral Prophylaxis
Deep Scaling
Fluoride Treatment
Orthodontic Treatment (Braces)
Dental Splint
Retainers
Night Guard
Tooth Extraction
Wisdom Tooth Extraction
Tooth Filling
Pit & Fissure Sealant
Esthetic Restoration
```

### Tab: `Config`
Column A = Key, Column B = Value:
| Key | Value |
|-----|-------|
| CLINIC_NAME | TomFord Dental |
| CLINIC_ADDRESS | RB & A BLDG., 166 Lakeview Drive, COR Kawilihan Lane, Pasig Blvd, Pasig, Philippines 1600 |
| CLINIC_PHONE | 0995 418 8879 |
| CLINIC_EMAIL | tomford.dental@gmail.com |
| CLINIC_HOURS | Mon–Sat  9:00 AM – 7:00 PM |
| CONCIERGE_EMAIL | tomford.dental@gmail.com |
| ADMIN_TOKEN | (set a secret password — used by the admin panel login) |
| ADMIN_URL | https://tomforddental.com/admin |
| CALENDAR_DURATION_MINS | 60 |
| BOOKING_TAGLINE | where every tooth matters. |
| OPEN_TIME | 09:00 |
| CLOSE_TIME | 19:00 |
| SLOT_DURATION_MINS | 30 |
| OPEN_DAYS | 1,2,3,4,5,6 |
| MAX_BOOKINGS_PER_SLOT | 1 |
| CALENDAR_ID | (your clinic Google Calendar ID) |
| SLOT_BUFFER_MINS | 60 |
| ADVANCE_BOOKING_DAYS | 30 |

> **CONCIERGE_EMAIL** — never sent to the frontend. Change this to update where clinic notification emails go.

> **ADMIN_TOKEN** — the password the clinic uses to log into the admin panel. Set something strong.

> **OPEN_DAYS** uses 0=Sunday, 1=Monday … 6=Saturday.

> **CALENDAR_ID** — set to the clinic's shared Google Calendar ID (not "primary"). Any event on that calendar blocks that time slot.

---

## 4. First-time Deploy as Web App

1. Click **Deploy → New deployment**
2. Type: **Web App**
3. Execute as: **Me**
4. Who has access: **Anyone**
5. Click **Deploy** → copy the **Web App URL**

## 5. Set SCRIPT_URL in Cloudflare (one-time)

The website never calls Apps Script directly — it goes through a Cloudflare proxy at `/api`.  
You only need to set the real URL **once** in Cloudflare:

1. Cloudflare Dashboard → **Workers & Pages** → `tomford-dental-website`
2. **Settings → Environment Variables → Add variable**
3. Name: `SCRIPT_URL`  
   Value: *(paste the Web App URL from step 4)*
4. Click **Save** — done. The frontend code never needs to change.

## 6. Grant Gmail permissions
On first run, Apps Script will ask for Gmail/Calendar permission. Click **Allow**.

## 7. Test
Submit a test booking on the website. You should:
- Patient receives a **"Request Received"** email
- Clinic receives a **"New Request"** notification email with a "Review in Admin" button
- Log in at `tomforddental.com/admin` → booking appears as **Pending**
- Click **Approve** → patient receives a **"Confirmed!"** email + calendar event created
- Click **Decline** → patient receives a polite rejection email

---

## Redeploying Apps Script (after code changes)

### ✅ The right way — URL stays the same forever:
1. Apps Script Editor → **Deploy → Manage Deployments**
2. Click the **pencil (Edit)** icon on your existing deployment
3. Change **Version** to **"New version"**
4. Click **Deploy** — same URL, updated code ✓

### ❌ Wrong way — creates a new URL every time:
> Deploy → **New deployment** — don't do this after the first time.  
> If you do end up with a new URL, just update `SCRIPT_URL` in Cloudflare Environment Variables. No code changes needed.

---

## Admin Panel

Live at: `https://tomforddental.com/admin`

- Login with the `ADMIN_TOKEN` you set in the Config sheet
- **Pending** tab — review and approve or decline requests
- **All Bookings** tab — full history table
- CRM features (Patients, Analytics, Reminders, Payments) — coming in a future update
