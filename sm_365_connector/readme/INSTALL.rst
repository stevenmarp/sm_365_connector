1. Place ``sm_365_connector`` folder in your Odoo addons path.
2. Restart Odoo server.
3. Go to **Apps**, update the Apps List, search for *Microsoft 365 Connector* and click **Install**.

Prerequisites
~~~~~~~~~~~~~

- An **Azure AD App Registration** with the following API permissions (Delegated):
  ``User.Read``, ``Contacts.ReadWrite``, ``Calendars.ReadWrite``, ``Mail.ReadWrite``,
  ``Mail.Send``, ``Tasks.ReadWrite``, ``Files.ReadWrite``, ``OnlineMeetings.ReadWrite``,
  ``offline_access``.
- A **Client Secret** created in Azure AD → Certificates & Secrets.
- For Dynamics 365 CRM: an **Application (Client Credentials)** permission for Dynamics CRM
  (``https://<org>.api.crm.dynamics.com/.default``).
- For webhooks: your Odoo instance must be accessible via a **public HTTPS** domain.

Dependencies
~~~~~~~~~~~~

This module depends on:

- ``base``
- ``mail``
- ``contacts``
- ``calendar``
- ``crm``
- ``project``
- ``auth_oauth``

.. important::
   The ``crm`` and ``project`` modules must be installed for Dynamics 365 CRM import
   and Microsoft To Do sync to function properly.
