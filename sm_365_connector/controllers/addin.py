# -*- coding: utf-8 -*-

import base64
import json
import logging

from odoo import http
from odoo.http import request

_logger = logging.getLogger(__name__)


class OutlookAddinController(http.Controller):
    """Serve the Outlook Web Add-in manifest and task-pane HTML.

    To install:
      1. Open Outlook on the web → Settings → General → Manage add-ins
      2. Add Custom Add-in → From URL
      3. Enter:  https://<your-odoo-url>/microsoft365/addin/manifest.xml
    """

    @http.route(
        '/microsoft365/addin/manifest.xml', type='http', auth='public',
        csrf=False, save_session=False,
    )
    def manifest(self, **kw):
        base_url = request.env['ir.config_parameter'].sudo().get_param('web.base.url')
        xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
           xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
           xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
           xmlns:mailappor="http://schemas.microsoft.com/office/mailappversionoverrides/1.0"
           xsi:type="MailApp">
  <Id>7b1f2e8c-0000-4d5f-9e00-365connector</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Steven Marp</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Odoo 365 Connector"/>
  <Description DefaultValue="Create Odoo contacts, leads &amp; tasks directly from Outlook."/>
  <IconUrl DefaultValue="{base_url}/microsoft365/addin/icon-64.png"/>
  <HighResolutionIconUrl DefaultValue="{base_url}/microsoft365/addin/icon-128.png"/>
  <SupportUrl DefaultValue="{base_url}"/>
  <AppDomains>
    <AppDomain>{base_url}</AppDomain>
  </AppDomains>
  <Hosts>
    <Host Name="Mailbox"/>
  </Hosts>
  <Requirements>
    <Sets>
      <Set Name="Mailbox" MinVersion="1.1"/>
    </Sets>
  </Requirements>
  <FormSettings>
    <Form xsi:type="ItemRead">
      <DesktopSettings>
        <SourceLocation DefaultValue="{base_url}/microsoft365/addin/taskpane"/>
        <RequestedHeight>300</RequestedHeight>
      </DesktopSettings>
    </Form>
  </FormSettings>
  <Permissions>ReadItem</Permissions>
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read"/>
  </Rule>
  <DisableEntityHighlighting>false</DisableEntityHighlighting>
</OfficeApp>"""
        return request.make_response(
            xml, headers=[('Content-Type', 'application/xml')]
        )

    @http.route(
        '/microsoft365/addin/taskpane', type='http', auth='public',
        csrf=False, save_session=False,
    )
    def taskpane(self, **kw):
        base_url = request.env['ir.config_parameter'].sudo().get_param('web.base.url')
        html = f"""<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8"/>
  <title>Odoo 365 Connector</title>
  <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
  <style>
    body {{ font-family: -apple-system, Segoe UI, Roboto, sans-serif; padding: 12px; }}
    h3 {{ margin: 0 0 10px 0; color: #714B67; }}
    button {{ background: #714B67; color: white; border: 0; padding: 8px 12px;
             border-radius: 4px; cursor: pointer; margin: 4px 0; width: 100%; }}
    button:hover {{ background: #5a3b52; }}
    .info {{ background: #f5f5f5; padding: 8px; border-radius: 4px; margin: 8px 0; font-size: 12px; }}
    .ok {{ color: green; }} .err {{ color: red; }}
  </style>
</head>
<body>
  <h3>Odoo 365 Connector</h3>
  <div id="emailInfo" class="info">Loading…</div>
  <button onclick="createLead()">Create CRM Lead</button>
  <button onclick="createContact()">Create Contact</button>
  <button onclick="createTask()">Create Task</button>
  <div id="result"></div>
  <script>
    const ODOO_URL = '{base_url}';
    let currentItem = {{}};
    Office.onReady(() => {{
      const item = Office.context.mailbox.item;
      currentItem = {{
        from: item.from && item.from.emailAddress,
        fromName: item.from && item.from.displayName,
        subject: item.subject,
        body: '',
      }};
      document.getElementById('emailInfo').innerHTML =
        '<b>From:</b> ' + currentItem.fromName + ' &lt;' + currentItem.from + '&gt;<br/>' +
        '<b>Subject:</b> ' + currentItem.subject;
    }});

    function postToOdoo(endpoint) {{
      fetch(ODOO_URL + endpoint, {{
        method: 'POST',
        headers: {{'Content-Type': 'application/json'}},
        credentials: 'include',
        body: JSON.stringify({{jsonrpc: '2.0', params: currentItem}}),
      }})
      .then(r => r.json())
      .then(d => {{
        const res = d.result || {{}};
        if (res.ok) {{
          document.getElementById('result').innerHTML =
            '<p class="ok">✓ ' + res.message + '</p>';
        }} else {{
          document.getElementById('result').innerHTML =
            '<p class="err">✗ ' + (res.message || 'Error') + '</p>';
        }}
      }})
      .catch(e => {{
        document.getElementById('result').innerHTML =
          '<p class="err">✗ ' + e.message + ' — please login to Odoo first.</p>';
      }});
    }}

    function createLead()    {{ postToOdoo('/microsoft365/addin/create_lead'); }}
    function createContact() {{ postToOdoo('/microsoft365/addin/create_contact'); }}
    function createTask()    {{ postToOdoo('/microsoft365/addin/create_task'); }}
  </script>
</body>
</html>"""
        return request.make_response(
            html, headers=[('Content-Type', 'text/html; charset=utf-8')]
        )

    def _respond(self, ok, message, record_id=None):
        return {'ok': ok, 'message': message, 'record_id': record_id}

    @http.route(
        '/microsoft365/addin/create_lead', type='json', auth='user',
        csrf=False, methods=['POST'],
    )
    def create_lead(self, **kw):
        params = request.params or {}
        email = params.get('from') or ''
        name = params.get('fromName') or email
        subject = params.get('subject') or 'From Outlook'
        lead = request.env['crm.lead'].create({
            'name': subject,
            'email_from': email,
            'contact_name': name,
            'type': 'lead',
        })
        return self._respond(True, f'Lead #{lead.id} created', lead.id)

    @http.route(
        '/microsoft365/addin/create_contact', type='json', auth='user',
        csrf=False, methods=['POST'],
    )
    def create_contact(self, **kw):
        params = request.params or {}
        email = params.get('from') or ''
        name = params.get('fromName') or email or 'Contact'
        partner = request.env['res.partner'].create({
            'name': name, 'email': email,
        })
        return self._respond(True, f'Contact #{partner.id} created', partner.id)

    @http.route(
        '/microsoft365/addin/create_task', type='json', auth='user',
        csrf=False, methods=['POST'],
    )
    def create_task(self, **kw):
        params = request.params or {}
        task = request.env['project.task'].create({
            'name': params.get('subject') or 'Task from Outlook',
            'description': 'From: %s' % (params.get('from') or ''),
        })
        return self._respond(True, f'Task #{task.id} created', task.id)
