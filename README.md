# A1 Dental Clinic

Website based on the approved GALADENT+ dental clinic architecture.

## Architecture
- `index.html` — public prices and online booking.
- `about.html` — clinic information.
- `tips.html` — patient content.
- `staff.html` — separate internal staff/doctor cabinet.
- `Code.gs` — Google Apps Script backend shared by the public booking page and staff cabinet.

## Per-clinic connections
A1 must use its own Google resources. Never reuse another clinic's backend URL, Sheet or Calendar.

1. Create a Google Sheet for A1 and put its ID into `SHEET_ID` in `Code.gs`.
2. Create a dedicated Google Calendar for A1 and put its ID into `CALENDAR_ID` in `Code.gs`.
3. Deploy `Code.gs` as a Google Apps Script Web App.
4. Put that Web App `/exec` URL into `WEB_APP_URL` in both `index.html` and `staff.html`.
5. After any Apps Script code change, deploy a new Web App version.

The backend creates/uses the `Bookings` sheet and supports online booking, manual staff booking, free slots, cancellation, rescheduling, visit editing, daily lists and patient history search.

Current clinic: A1, 15 Zakaria Paliashvili St, Tbilisi.
