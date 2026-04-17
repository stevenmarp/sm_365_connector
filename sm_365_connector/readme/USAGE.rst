Outlook Contact Sync
~~~~~~~~~~~~~~~~~~~~~

1. Open a connection and click **Sync Contacts**.
2. Contacts from Outlook are imported/updated in Odoo (matched by email).
3. Odoo contacts with changes are pushed back to Outlook.
4. Stat button shows synced contact count.

Calendar Sync & Teams Meetings
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

1. Click **Sync Calendar** to pull/push events between Outlook and Odoo.
2. To create a Teams meeting, open any calendar event and click **Create Teams Meeting**.
3. The Teams join URL is automatically added to the event description.

Email Import & Send
~~~~~~~~~~~~~~~~~~~~

1. Click **Import Emails** to fetch recent emails from your Outlook inbox.
2. Emails are stored as ``mail.message`` records with full HTML body and attachments.
3. View all imported emails in the **Outlook Inbox** menu.
4. To send an email via Outlook, use the **Send via Outlook** button on any contact.
5. Emails are sent using the Microsoft Graph API — no SMTP configuration needed.

Email Auto-Conversion Rules
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

1. Define rules in the connection's **Email Rules** tab.
2. When emails are imported, each email is matched against rules (by sequence order).
3. Matching emails automatically create the configured record type (lead, opportunity, task, note).
4. Rules support exact text match or regex patterns.

Microsoft To Do Sync
~~~~~~~~~~~~~~~~~~~~~

1. Click **Sync To Do Tasks** to sync tasks between Microsoft To Do and Odoo Project.
2. New To Do tasks appear as Odoo project tasks, and vice versa.
3. Task titles, descriptions, due dates, and completion status stay in sync.

OneDrive Attachments
~~~~~~~~~~~~~~~~~~~~~

1. Enable OneDrive in the connection settings.
2. On any record with attachments, click **Upload to OneDrive**.
3. Files are uploaded to the configured OneDrive folder.
4. Direct download links and web view URLs are stored on each attachment.

Dynamics 365 CRM Import
~~~~~~~~~~~~~~~~~~~~~~~~~

1. Enable Dynamics 365 CRM and connect (see Configuration).
2. Use the import buttons:

   - **Import Leads** — imports active leads from Dynamics 365
   - **Import Opportunities** — imports open opportunities
   - **Import Accounts** — imports account records as Odoo company contacts
   - **Import Contacts** — imports contact records as Odoo person contacts

3. Records are deduplicated by their Dynamics 365 ID on re-import.
4. Source stage, estimated revenue, and other fields are mapped automatically.

Real-Time Webhooks
~~~~~~~~~~~~~~~~~~~

1. Enable webhooks in connection settings.
2. Microsoft Graph sends change notifications when contacts, events, or emails change.
3. Odoo processes these notifications and syncs the affected records instantly.
4. Webhook subscriptions auto-renew before expiry.

Outlook Web Add-in
~~~~~~~~~~~~~~~~~~~~

1. Download the Add-in manifest from the connection form.
2. Sideload the manifest in Outlook (Settings → Get Add-ins → My Add-ins → Custom).
3. When reading an email in Outlook, open the add-in panel.
4. Create contacts, leads, or tasks in Odoo directly from the email context.

Sync Logging
~~~~~~~~~~~~~

- Every sync operation is logged in **Microsoft 365 → Sync Logs**.
- Track status: success, partial, or error.
- View detailed error messages for failed operations.
- Filter by type (contacts, calendar, email, CRM) and date.

Technical Flow
~~~~~~~~~~~~~~~

1. Outlook auth uses OAuth2 Authorization Code flow with token refresh.
2. Dynamics auth uses OAuth2 Client Credentials flow.
3. All API calls go through ``_call_graph_api()`` and ``_call_dynamics_api()`` helpers.
4. Token refresh is automatic — if a token is expired, it is refreshed before the API call.
5. Webhook endpoint at ``/microsoft365/webhook`` validates client state and processes notifications.
6. Cron jobs handle periodic sync (every 60 min) and subscription renewal (every 6 hours).
