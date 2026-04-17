# -*- coding: utf-8 -*-
{
    'name': 'Microsoft 365 Connector — Outlook, Teams, OneDrive & Dynamics CRM',
    'version': '16.0.1.0.0',
    'summary': 'All-in-one Microsoft 365: 2-way Outlook sync (contacts, calendar, emails), Teams meetings, To Do tasks, OneDrive storage, email rules, real-time webhooks, Outlook add-in & Dynamics 365 CRM import. Direct API, no middleware.',
    'description': """
Microsoft 365 Connector for Odoo 16
=====================================

All-in-one connector to Microsoft 365 — Outlook, Teams, To Do, OneDrive & Dynamics CRM in a single module:

**Outlook Integration (via Microsoft Graph API):**
- 2-way contact sync (Outlook ↔ Odoo contacts)
- 2-way calendar sync (Outlook ↔ Odoo calendar)
- Import emails from Outlook inbox (HTML body, attachments, inline images)
- Send emails via Outlook Graph API (no SMTP needed)
- Unified Outlook inbox view inside Odoo
- Email auto-conversion rules (incoming emails → leads, opportunities, tasks, notes)
- Auto-sync via scheduled actions or real-time webhooks

**Microsoft Teams:**
- Create Teams meetings directly from any calendar event
- Auto-insert join URL, conference ID & dial-in info

**Microsoft To Do:**
- 2-way task sync (To Do ↔ Odoo Project tasks)
- Sync titles, descriptions, due dates & completion status

**OneDrive:**
- Offload Odoo attachments to OneDrive
- Direct download & web view links stored on attachments
- Configurable folder path

**Dynamics 365 CRM Integration (via Dynamics Web API):**
- Import Leads → Odoo CRM leads
- Import Opportunities → Odoo CRM opportunities
- Import Accounts → Odoo contacts (companies)
- Import Contacts → Odoo contacts (persons)
- Deduplication & update on re-import

**Real-Time Webhooks:**
- Microsoft Graph change notifications for instant sync
- Auto-renewing subscriptions via cron

**Outlook Web Add-in:**
- Built-in add-in manifest for sideloading
- Create contacts, leads & tasks from inside Outlook

**Technical Highlights:**
- Direct Microsoft Graph API & Dynamics Web API (no middleware, no third-party libraries)
- OAuth2 Authorization Code flow for Outlook
- OAuth2 Client Credentials flow for Dynamics CRM
- Automatic token refresh
- Full sync logging & history with error details
- Kanban dashboard with sync stats & quick actions
- Role-based access control (User / Manager)
- Multi-company support
    """,
    'author': 'Steven Marp',
    'website': 'https://apps.odoo.com/apps/browse?repo_maintainer_id=512936',
    'category': 'Productivity',
    'license': 'OPL-1',

    'depends': [
        'base',
        'mail',
        'contacts',
        'calendar',
        'crm',
        'project',
        'auth_oauth',
    ],

    'data': [
        # Security
        'security/security.xml',
        'security/ir.model.access.csv',

        # Wizard (must be before views that reference its action)
        'wizard/outlook_send_email_views.xml',

        # Views
        'views/microsoft365_connection_views.xml',
        'views/microsoft365_sync_log_views.xml',
        'views/res_partner_views.xml',
        'views/calendar_event_views.xml',
        'views/project_task_views.xml',
        'views/sm_365_email_rule_views.xml',
        'views/sm_365_inbox_views.xml',
        'views/sm_365_dashboard_views.xml',
        'views/menu.xml',

        # Data
        'data/cron.xml',
        'data/oauth_provider.xml',
    ],

    'assets': {
        'web.assets_backend': [
            'sm_365_connector/static/src/js/**/*',
            'sm_365_connector/static/src/scss/**/*',
        ],
    },

    'images': ['static/description/banner.gif'],
    'price': 1040.2,
    'currency': 'USD',
    'installable': True,
    'application': True,
    'auto_install': False,
}
