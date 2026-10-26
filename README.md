# Brampton Shipping Booking Form

This is a standalone Google Apps Script project for the **Brampton location only**.

## Workflow

1. A customer opens the standard web-app URL and submits a Brampton booking.
2. The booking is stored in the `Bookings` sheet as `Awaiting Brampton Office`.
3. Brampton staff open the same URL with `?view=office`, enter the office password, and access the dashboard.
4. The dashboard shows total, pending, and completed counters plus a tabular booking list.
5. Staff can filter by booking date, free-text search, or status.
6. The **Edit** button opens the full booking and the Brampton office processing fields.
7. Saving those fields marks the booking as `Completed`.
8. **Logout** clears the password and loaded records from the browser.

There is no location selector, access for other offices, or Master Admin role in this project.

## Office password

The office password is not stored in the source code. Configure it from the
Google Sheet using **Shipping Form → Configure project / change password**.
Only its SHA-256 hash is stored in Script Properties.

## Install in a new Google Sheet

1. Create a new Google Sheet for Brampton bookings.
2. Open **Extensions → Apps Script**.
3. Replace `Code.gs` with this project's `Code.gs`.
4. Add an HTML file named `Index` and paste this project's `Index.html`.
5. In Apps Script Project Settings, enable the manifest file and replace it with `appsscript.json`.
6. Save the project and reload the Google Sheet.
7. Select **Shipping Form → Configure project / change password** and enter a private password containing at least 10 characters.
8. Approve the requested Google permissions.

This creates and formats the `Bookings` sheet and activates the default password.

## Deploy

1. In Apps Script, select **Deploy → New deployment**.
2. Choose **Web app**.
3. Set **Execute as** to **Me**.
4. Choose the narrowest access option that still allows intended customers.
5. Deploy and copy the `/exec` URL.

Use:

```text
Customer form:  https://script.google.com/macros/s/DEPLOYMENT_ID/exec
Brampton office: https://script.google.com/macros/s/DEPLOYMENT_ID/exec?view=office
```

After changing `Code.gs` or `Index.html`, create a new deployment version. A password-only change applied through the Sheet menu does not require redeployment.

## GitHub Pages address

The lowercase `index.html` is a GitHub Pages entry page. It forwards query
parameters to the Apps Script deployment, so both addresses work:

```text
Customer form:  https://quickship-now.github.io/brampton-shipping-booking-form/
Brampton office: https://quickship-now.github.io/brampton-shipping-booking-form/?view=office
```

The browser will ultimately display the Google Apps Script address because the
application uses `google.script.run`, which is available only on Apps Script
HTML-service pages.

## Files

- `Code.gs` — Brampton-only authentication, validation, Google Sheets storage, and office processing.
- `Index.html` — Brampton customer form and filtered office dashboard.
- `index.html` — GitHub Pages entry page that redirects to the Apps Script deployment.
- `appsscript.json` — Apps Script runtime and permissions.

