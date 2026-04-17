====================================================
Microsoft 365 Connector — User Guide
====================================================

.. contents:: Table of Contents
   :depth: 2
   :local:

----

1. Overview
===========

The **Microsoft 365 Connector** brings your entire Microsoft 365 ecosystem into Odoo:
Outlook emails, contacts, calendar, Teams meetings, To Do tasks, OneDrive files,
and Dynamics 365 CRM — all from a single module.

Everything syncs automatically. No middleware, no monthly subscription fees.
Install once, connect, and work from Odoo without switching between apps.

----

2. What You Need Before Starting
=================================

Before you begin, make sure you have:

- A **Microsoft 365 account** (personal Outlook works for testing, but
  Microsoft 365 Business is recommended for production)
- Access to **Azure Portal** (portal.azure.com) — this is where you register your app
- Your Odoo server accessible via **HTTPS** (required by Microsoft for security)
- For Dynamics 365 CRM features: a **Dynamics 365 license** on your tenant

.. tip::
   Don't have HTTPS yet? For local testing, use a tunneling service like
   `ngrok <https://ngrok.com>`_ to temporarily expose your Odoo over HTTPS.

----

3. Setting Up Azure (Step by Step)
===================================

This is a one-time setup. Takes about 10 minutes.

Step 1 — Register Your App
----------------------------

1. Go to **portal.azure.com** and sign in.
2. Search for **"App registrations"** in the top search bar.
3. Click **New registration**.
4. Fill in:

   - **Name**: anything you like, e.g. "Odoo Microsoft 365"
   - **Supported account types**: choose "All Microsoft account users" (recommended)
   - **Redirect URI**: select **Web** and enter:
     ``https://your-odoo-domain.com/microsoft365/callback``

5. Click **Register**. Done!

.. important::
   The Redirect URI must **exactly match** your Odoo domain with HTTPS.
   If your Odoo is at ``https://erp.mycompany.com``, the URI would be
   ``https://erp.mycompany.com/microsoft365/callback``.

Step 2 — Grant Permissions
---------------------------

Your app needs permission to access Outlook, Calendar, etc.

1. In your newly created app, click **API permissions** in the left menu.
2. Click **Add a permission** → **Microsoft Graph** → **Delegated permissions**.
3. Search and add these permissions:

   - ``User.Read``
   - ``Contacts.ReadWrite``
   - ``Calendars.ReadWrite``
   - ``Mail.ReadWrite``
   - ``Mail.Send``
   - ``Tasks.ReadWrite``
   - ``Files.ReadWrite``
   - ``OnlineMeetings.ReadWrite``
   - ``offline_access``

4. Click **Grant admin consent** (the green checkmark button at the top).

Step 3 — Create a Secret Key
------------------------------

1. In the left menu, click **Certificates & secrets**.
2. Click **New client secret**.
3. Give it a name (e.g. "Odoo") and choose an expiry (24 months recommended).
4. Click **Add**.
5. **Copy the Value immediately** — you won't be able to see it again!

Step 4 — Copy Your App Credentials
-------------------------------------

Go back to the **Overview** page of your app. You need three values:

+----------------------------+----------------------------------------------+
| What to copy               | Where to find it                             |
+============================+==============================================+
| **Application (Client) ID** | Overview page, first line                   |
+----------------------------+----------------------------------------------+
| **Directory (Tenant) ID**  | Overview page, second line                   |
+----------------------------+----------------------------------------------+
| **Client Secret**          | The value you copied in Step 3               |
+----------------------------+----------------------------------------------+

.. warning::
   Keep your Client Secret safe! Don't share it publicly or commit it to version control.

----

4. Connecting Odoo to Microsoft 365
=====================================

Now switch to Odoo.

Step 1 — Create a Connection
------------------------------

1. Open the **Microsoft 365** menu (appears after installing the module).
2. Click **Connections** → **Create**.
3. Fill in the three credentials from Azure:

   - **Tenant ID** (or type ``common`` if you chose multi-tenant in Azure)
   - **Application (Client) ID**
   - **Client Secret**

4. Give it a name, e.g. "My Company 365".

Step 2 — Authorize Outlook
----------------------------

1. Click the **Authorize Outlook** button.
2. You'll be redirected to Microsoft's login page.
3. Sign in with your Microsoft account.
4. Accept the permission request.
5. Microsoft redirects you back to Odoo — the status changes to **Outlook Connected**.

Step 3 — Verify It Works
--------------------------

1. Click **Test Connection**.
2. You should see a success message with your Microsoft account name and email.
3. You're connected!

That's it for basic setup. All Outlook features (contacts, calendar, email) are now ready to use.

----

5. Setting Up Dynamics 365 CRM (Optional)
==========================================

Skip this if you don't use Dynamics 365 CRM.

.. note::
   Dynamics 365 requires its own license, separate from regular Microsoft 365.
   The same Azure app can be used for both Outlook and Dynamics.

1. In Azure Portal → your App Registration → **API Permissions**.
2. Click **Add a permission** → **Dynamics CRM** → **Application permissions**.
3. Add ``user_impersonation`` and grant admin consent.
4. In Odoo, open your connection and enable **Enable Dynamics 365 CRM**.
5. Enter your **Organization Name** — this is the subdomain from your Dynamics URL:

   - ``https://contoso.crm.dynamics.com`` → Organization Name is ``contoso``
   - ``https://mycompany.crm4.dynamics.com`` → Organization Name is ``mycompany``

6. Select the correct **Region** for your Dynamics environment.
7. Click **Connect Dynamics**.

.. warning::
   The Organization Name is **not** your company display name. It's the short
   identifier in your Dynamics URL. Using a display name with spaces will fail.

----

6. Using the Module
===================

Contact Sync
-------------

- Click **Sync Contacts** to sync Outlook ↔ Odoo contacts in both directions.
- Contacts are matched by email address to prevent duplicates.
- New contacts on either side are automatically created on the other.

Calendar Sync
--------------

- Click **Sync Calendar** to sync events between Outlook and Odoo.
- Changes on either side are reflected on the other after sync.

Teams Meetings
---------------

- Open any calendar event in Odoo.
- Click **Create Teams Meeting** — the join link is added automatically.
- Attendees can click the link to join the meeting.
- Requires a Microsoft Teams license on the connected account.

Email
------

**Importing emails:**

- Click **Import Emails** to pull recent emails from your Outlook inbox into Odoo.
- Emails include the full HTML body, attachments, and images.
- View them in **Microsoft 365 → Outlook Inbox**.

**Sending emails:**

- On any contact, click **Send via Outlook**.
- Write your message and click Send — delivered via your Outlook account.
- No SMTP setup needed.

**Auto-conversion rules:**

- Define rules in the **Email Rules** tab on your connection.
- Incoming emails matching your rules automatically create leads, opportunities,
  tasks, or internal notes.

To Do Tasks
------------

- Click **Sync To Do Tasks** to sync between Microsoft To Do and Odoo Project.
- Tasks sync in both directions: titles, descriptions, due dates, and status.

OneDrive
---------

- Enable **Use OneDrive for Attachments** in your connection settings.
- Click **Upload to OneDrive** on any record to offload attachments.
- Files get direct download and view links stored in Odoo.

Dynamics 365 CRM
------------------

Use the import buttons on your connection:

- **Import Leads** — brings active leads into Odoo CRM
- **Import Opportunities** — brings open opportunities
- **Import Accounts** — imports companies as Odoo contacts
- **Import Contacts** — imports people as Odoo contacts

Re-importing updates existing records instead of creating duplicates.

----

7. Real-Time Sync (Webhooks)
=============================

Instead of manual or scheduled sync, you can enable **instant sync**:

1. Enable **Real-Time Webhooks** in your connection settings.
2. Click **Enable Webhooks**.
3. Now whenever a contact, event, or email changes in Outlook,
   Odoo is notified and syncs immediately.

Subscriptions are automatically renewed — no maintenance needed.

.. important::
   Webhooks require your Odoo server to be accessible from the internet
   via HTTPS. Self-signed certificates won't work.

----

8. Outlook Add-in
==================

Work with Odoo without leaving Outlook:

1. Download the add-in manifest from your connection form.
2. In Outlook → **Settings** → **Get Add-ins** → **My Add-ins** → upload the manifest.
3. When reading an email, open the add-in panel.
4. Create contacts, leads, or tasks in Odoo directly from the email.

----

9. Automatic Sync Schedule
============================

The module includes scheduled jobs that run automatically:

- **Auto-Sync** — syncs contacts, calendar, emails, and tasks periodically
  (default: every 60 minutes). Enable **Scheduled Auto-Sync** in connection settings.
- **Webhook Renewal** — renews real-time sync subscriptions before they expire
  (runs every 6 hours automatically).

----

10. Single Sign-On (SSO)
=========================

When you connect Outlook, the module automatically enables **"Sign in with Microsoft 365"**
on your Odoo login page. Your team can log into Odoo using their Microsoft credentials
without a separate password.

SSO is automatically disabled when all Microsoft 365 connections are disconnected.

----

11. Troubleshooting
====================

**"redirect_uri is not valid" error when authorizing**

→ The Redirect URI in Azure doesn't match your Odoo URL. Make sure you added
``https://your-domain.com/microsoft365/callback`` in Azure → Authentication.
Must be HTTPS, must match exactly.

**OAuth login shows "AADSTS" error**

→ Double-check the Tenant ID in Odoo. Try ``common`` for multi-tenant support.

**Token refresh fails / "re-authorize" message**

→ Your Client Secret may have expired. Create a new one in Azure and update
it in the Odoo connection form.

**Dynamics connection fails**

→ Make sure the Organization Name is the URL subdomain (e.g. ``contoso``),
not your company display name. Check that the Region matches your environment.

**Webhooks not working**

→ Your Odoo server must be publicly accessible via HTTPS.
Self-signed certificates are rejected by Microsoft.

**Contact sync creates duplicates**

→ Contacts are matched by email address. Make sure both sides have the same
email for proper deduplication.

**Teams meeting creation fails**

→ The connected Microsoft account needs a Teams license.

**Emails rejected / go to spam**

→ Free Outlook accounts (``@outlook.com``) have low sender reputation.
For production use, we recommend Microsoft 365 Business with a custom domain.

**Email images broken after import**

→ Click **Re-fetch Email Images** on the connection to re-download inline images.

----

12. Security
============

- All communication with Microsoft uses **OAuth 2.0** — no passwords are stored.
- Tokens are encrypted and only accessible to system administrators.
- Tokens refresh automatically — no manual maintenance.
- Webhook notifications are validated for authenticity.
- Two security groups: **365 User** (basic access) and **365 Manager** (full admin).
- All data transfer happens over encrypted HTTPS connections.