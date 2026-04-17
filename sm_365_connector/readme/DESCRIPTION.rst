Microsoft 365 Connector for Odoo 19 — all-in-one integration to Microsoft 365
combining Outlook and Dynamics 365 CRM in a single module. Direct API calls to
Microsoft Graph and Dynamics Web API, no middleware, no third-party libraries.

**Outlook Integration (via Microsoft Graph API):**

- 2-way contact sync (Outlook ↔ Odoo contacts)
- 2-way calendar sync (Outlook ↔ Odoo calendar)
- Import emails from Outlook inbox with full HTML and attachments
- Send emails via Outlook using Graph API (no SMTP)
- Auto-conversion rules: incoming emails → leads, opportunities, tasks, notes
- Create Microsoft Teams meetings from any calendar event
- Sync tasks with Microsoft To Do
- Store attachments on OneDrive
- Real-time webhooks for instant sync
- Outlook Web Add-in for record creation inside Outlook

**Dynamics 365 CRM Integration (via Dynamics Web API):**

- Import Leads → Odoo CRM leads
- Import Opportunities → Odoo CRM opportunities
- Import Accounts → Odoo contacts (companies)
- Import Contacts → Odoo contacts (persons)
- Deduplication and update on re-import

Built for businesses using Microsoft 365 who want all their data accessible
from inside Odoo without switching between applications.
