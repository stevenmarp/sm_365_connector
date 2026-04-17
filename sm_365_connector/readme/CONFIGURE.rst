Step 1: Register an App in Azure
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

1. Go to **portal.azure.com** → search **"App registrations"** → **New registration**.
2. Name it (e.g. "Odoo Microsoft 365"), select "All Microsoft account users".
3. Set Redirect URI to **Web**: ``https://your-odoo-domain.com/microsoft365/callback``
4. Click **Register**.

Step 2: Grant Permissions
~~~~~~~~~~~~~~~~~~~~~~~~~~

1. In your app → **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated**.
2. Add: ``User.Read``, ``Contacts.ReadWrite``, ``Calendars.ReadWrite``,
   ``Mail.ReadWrite``, ``Mail.Send``, ``Tasks.ReadWrite``, ``Files.ReadWrite``,
   ``OnlineMeetings.ReadWrite``, ``offline_access``.
3. Click **Grant admin consent**.

Step 3: Create a Client Secret
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

1. Go to **Certificates & secrets** → **New client secret** → set 24 months expiry.
2. **Copy the Value immediately** — it won't be shown again.

Step 4: Copy Your Credentials
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

From the app's **Overview** page, note:

- **Application (Client) ID**
- **Directory (Tenant) ID**
- **Client Secret** (from Step 3)

Step 5: Connect in Odoo
~~~~~~~~~~~~~~~~~~~~~~~~~

1. Install this module → go to **Microsoft 365** menu → **Connections** → **Create**.
2. Fill in Tenant ID, Client ID, and Client Secret.
3. Click **Authorize Outlook** → sign in with your Microsoft account.
4. Click **Test Connection** to verify.

Step 6: Dynamics 365 CRM (Optional)
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

1. In Azure → your app → **API Permissions** → add **Dynamics CRM** → ``user_impersonation``.
2. In Odoo, enable **Dynamics 365 CRM** on your connection.
3. Enter your **Organization Name** (the subdomain from your Dynamics URL, e.g. ``contoso``).
4. Select **Region** and click **Connect Dynamics**.

Step 7: Enable Webhooks (Optional)
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

For real-time sync instead of periodic polling:

1. Enable **Real-Time Webhooks** on your connection.
2. Click **Enable Webhooks** — Odoo is notified instantly when data changes in Outlook.
3. Requires HTTPS. Subscriptions auto-renew automatically.