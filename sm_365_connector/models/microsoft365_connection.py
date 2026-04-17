# -*- coding: utf-8 -*-

import base64
import json
import logging
import urllib.parse
from datetime import datetime, timedelta

import requests

from odoo import api, fields, models, _
from odoo.exceptions import UserError

_logger = logging.getLogger(__name__)

GRAPH_API_URL = 'https://graph.microsoft.com/v1.0'
GRAPH_AUTH_URL = 'https://login.microsoftonline.com/{tenant}/oauth2/v2.0'
GRAPH_SCOPES = (
    'offline_access User.Read Contacts.ReadWrite '
    'Calendars.ReadWrite Mail.ReadWrite Mail.Send '
    'Tasks.ReadWrite Files.ReadWrite OnlineMeetings.ReadWrite'
)


class Microsoft365Connection(models.Model):
    _name = 'sm.365.connection'
    _inherit = ['mail.thread', 'mail.activity.mixin']
    _description = 'Microsoft 365 Connection'
    _order = 'id desc'

    name = fields.Char(required=True, tracking=True)
    company_id = fields.Many2one(
        'res.company', default=lambda self: self.env.company, required=True,
    )
    user_id = fields.Many2one(
        'res.users', string='Owner User',
        help='If set, only this user can use this connection. Leave blank for a shared/company-wide connection.',
    )
    is_shared = fields.Boolean(
        'Shared Connection', default=True,
        help='Allow all users to use this connection. Uncheck for per-user personal connections.',
    )

    # ── Azure AD App ──
    tenant_id = fields.Char('Tenant ID', default='common', required=True)
    client_id = fields.Char('Application (Client) ID', required=True)
    client_secret = fields.Char('Client Secret', required=True)

    # ── Outlook OAuth2 tokens ──
    outlook_access_token = fields.Char(groups='base.group_system')
    outlook_refresh_token = fields.Char(groups='base.group_system')
    outlook_token_expiry = fields.Datetime()
    outlook_user_email = fields.Char('Connected Outlook Account', readonly=True)
    is_outlook_connected = fields.Boolean(compute='_compute_is_connected', store=False)

    # ── OneDrive backend ──
    onedrive_enabled = fields.Boolean('Use OneDrive for Attachments')
    onedrive_folder_path = fields.Char(
        'OneDrive Folder', default='/Apps/Odoo',
        help='Root folder on OneDrive where Odoo attachments are stored.',
    )

    # ── Email rules ──
    email_rule_ids = fields.One2many('sm.365.email.rule', 'connection_id')
    email_rule_count = fields.Integer(compute='_compute_counts')

    # ── Webhook subscriptions ──
    subscription_ids = fields.One2many('sm.365.subscription', 'connection_id')
    subscription_count = fields.Integer(compute='_compute_counts')
    webhooks_enabled = fields.Boolean(
        'Enable Real-Time Webhooks',
        help='Use Microsoft Graph change notifications for instant sync. Requires a publicly reachable HTTPS URL.',
    )

    # ── Teams meeting provider email (for Teams) ──
    teams_default_organizer = fields.Char(
        'Default Teams Organizer Email', readonly=True,
        help='Email of the user who will host Teams meetings. Defaults to the connected Outlook account.',
    )

    # ── Dynamics 365 CRM ──
    dynamics_enabled = fields.Boolean('Enable Dynamics 365 CRM')
    dynamics_org = fields.Char('Organization Name')
    dynamics_region = fields.Selection([
        ('crm', 'North America'),
        ('crm2', 'South America'),
        ('crm4', 'EMEA'),
        ('crm5', 'APAC'),
        ('crm6', 'Oceania'),
        ('crm7', 'Japan'),
        ('crm8', 'India'),
        ('crm9', 'North America 2'),
        ('crm11', 'United Kingdom'),
    ], default='crm', string='Region')
    dynamics_access_token = fields.Char(groups='base.group_system')
    dynamics_token_expiry = fields.Datetime()
    is_dynamics_connected = fields.Boolean(compute='_compute_is_connected', store=False)

    # ── Sync options ──
    sync_contacts = fields.Boolean('Sync Contacts', default=True)
    sync_calendar = fields.Boolean('Sync Calendar', default=True)
    sync_emails = fields.Boolean('Import Emails', default=True)
    sync_crm = fields.Boolean('Import CRM Data')
    auto_sync = fields.Boolean('Scheduled Auto-Sync')
    sync_interval = fields.Integer('Interval (minutes)', default=60)
    last_sync = fields.Datetime('Last Sync', readonly=True)

    state = fields.Selection([
        ('draft', 'Not Connected'),
        ('outlook_connected', 'Outlook Connected'),
        ('fully_connected', 'Fully Connected'),
        ('error', 'Error'),
    ], default='draft', tracking=True, readonly=True)

    sync_log_ids = fields.One2many('sm.365.sync.log', 'connection_id')
    sync_log_count = fields.Integer(compute='_compute_counts')
    sync_log_error_count = fields.Integer(compute='_compute_counts')
    contact_synced_count = fields.Integer(compute='_compute_counts')
    event_synced_count = fields.Integer(compute='_compute_counts')
    email_synced_count = fields.Integer(compute='_compute_counts')
    task_synced_count = fields.Integer(compute='_compute_counts')

    @api.depends('sync_log_ids', 'email_rule_ids', 'subscription_ids')
    def _compute_counts(self):
        Log = self.env['sm.365.sync.log']
        Partner = self.env['res.partner']
        Event = self.env['calendar.event']
        MailMsg = self.env['mail.message']
        Task = self.env['project.task']
        for rec in self:
            rec.sync_log_count = Log.search_count(
                [('connection_id', '=', rec.id)]
            )
            rec.sync_log_error_count = Log.search_count([
                ('connection_id', '=', rec.id), ('status', '=', 'error'),
            ])
            rec.email_rule_count = len(rec.email_rule_ids)
            rec.subscription_count = len(rec.subscription_ids)
            rec.contact_synced_count = Partner.search_count(
                [('microsoft365_id', '!=', False)]
            )
            rec.event_synced_count = Event.search_count(
                [('microsoft365_id', '!=', False)]
            )
            rec.email_synced_count = MailMsg.search_count(
                [('microsoft365_id', '!=', False)]
            )
            rec.task_synced_count = Task.search_count(
                [('microsoft365_id', '!=', False)]
            )

    def _compute_is_connected(self):
        for rec in self:
            rec.is_outlook_connected = bool(rec.sudo().outlook_access_token)
            rec.is_dynamics_connected = bool(rec.sudo().dynamics_access_token)

    @api.model
    def _get_user_connection(self):
        """Return the best connection for the current user."""
        user = self.env.user
        # Prefer a personal connection, fall back to shared company connection.
        conn = self.search([
            ('user_id', '=', user.id),
            ('state', 'in', ('outlook_connected', 'fully_connected')),
        ], limit=1)
        if conn:
            return conn
        return self.search([
            '|', ('is_shared', '=', True), ('user_id', '=', False),
            ('state', 'in', ('outlook_connected', 'fully_connected')),
            ('company_id', 'in', (False, user.company_id.id)),
        ], limit=1)

    def _get_redirect_uri(self):
        base_url = self.env['ir.config_parameter'].sudo().get_param('web.base.url')
        # Force HTTPS — Microsoft rejects http:// redirect URIs
        if base_url and base_url.startswith('http://'):
            base_url = 'https://' + base_url[7:]
        return f'{base_url}/microsoft365/callback'

    # ================================================================
    #  OUTLOOK AUTH  (Authorization Code + PKCE)
    # ================================================================

    def action_authorize_outlook(self):
        self.ensure_one()
        state = base64.urlsafe_b64encode(json.dumps({
            'cid': self.id,
            'db': self.env.cr.dbname,
        }).encode()).decode()

        params = {
            'client_id': self.client_id,
            'response_type': 'code',
            'redirect_uri': self._get_redirect_uri(),
            'response_mode': 'query',
            'scope': GRAPH_SCOPES,
            'state': state,
        }
        auth_url = (
            GRAPH_AUTH_URL.format(tenant=self.tenant_id)
            + '/authorize?'
            + urllib.parse.urlencode(params)
        )
        return {
            'type': 'ir.actions.act_url',
            'url': auth_url,
            'target': 'self',
        }

    def _exchange_code_for_token(self, code):
        """Called by the OAuth2 callback controller."""
        url = GRAPH_AUTH_URL.format(tenant=self.tenant_id) + '/token'
        data = {
            'client_id': self.client_id,
            'client_secret': self.client_secret,
            'code': code,
            'redirect_uri': self._get_redirect_uri(),
            'grant_type': 'authorization_code',
            'scope': GRAPH_SCOPES,
        }
        resp = requests.post(url, data=data, timeout=30)
        if resp.status_code != 200:
            self.write({'state': 'error'})
            raise UserError(_('Token exchange failed: %s') % resp.text)

        tokens = resp.json()
        vals = {
            'outlook_access_token': tokens['access_token'],
            'outlook_refresh_token': tokens.get('refresh_token'),
            'outlook_token_expiry': (
                fields.Datetime.now()
                + timedelta(seconds=tokens.get('expires_in', 3600))
            ),
            'state': (
                'fully_connected'
                if self.dynamics_enabled and self.dynamics_access_token
                else 'outlook_connected'
            ),
        }
        self.write(vals)

        # Fetch connected user info
        try:
            user_info = self._call_graph_api('get', '/me')
            self.outlook_user_email = user_info.get(
                'mail', user_info.get('userPrincipalName', '')
            )
        except Exception:
            pass

        # Auto-configure SSO provider with this connection's Client ID
        self._configure_sso_provider()

        self._create_sync_log('outlook_auth', 'success', 'Outlook connected')

    def _refresh_outlook_token(self):
        if not self.outlook_refresh_token:
            raise UserError(_('No refresh token. Please re-authorize Outlook.'))

        url = GRAPH_AUTH_URL.format(tenant=self.tenant_id) + '/token'
        data = {
            'client_id': self.client_id,
            'client_secret': self.client_secret,
            'refresh_token': self.outlook_refresh_token,
            'grant_type': 'refresh_token',
            'scope': GRAPH_SCOPES,
        }
        resp = requests.post(url, data=data, timeout=30)
        if resp.status_code != 200:
            self.write({'state': 'error'})
            raise UserError(_('Token refresh failed: %s') % resp.text)

        tokens = resp.json()
        self.write({
            'outlook_access_token': tokens['access_token'],
            'outlook_refresh_token': tokens.get(
                'refresh_token', self.outlook_refresh_token
            ),
            'outlook_token_expiry': (
                fields.Datetime.now()
                + timedelta(seconds=tokens.get('expires_in', 3600))
            ),
        })

    def _ensure_outlook_token(self):
        self.ensure_one()
        if not self.outlook_access_token:
            raise UserError(_('Not connected to Outlook. Click "Authorize Outlook".'))
        if self.outlook_token_expiry and self.outlook_token_expiry <= fields.Datetime.now():
            self._refresh_outlook_token()

    def _call_graph_api(self, method, endpoint, data=None, params=None):
        self._ensure_outlook_token()
        headers = {
            'Authorization': f'Bearer {self.outlook_access_token}',
            'Content-Type': 'application/json',
        }
        url = f'{GRAPH_API_URL}{endpoint}'
        try:
            resp = getattr(requests, method)(
                url, headers=headers, json=data, params=params, timeout=30
            )
            resp.raise_for_status()
            return resp.json() if resp.content else {}
        except requests.exceptions.HTTPError:
            _logger.error('Graph API %s %s → %s', method.upper(), endpoint, resp.text)
            raise UserError(_('Microsoft Graph API error: %s') % resp.text)
        except requests.exceptions.ConnectionError as exc:
            raise UserError(_('Cannot reach Microsoft servers: %s') % exc)

    # ================================================================
    #  DYNAMICS 365 CRM AUTH  (Client Credentials)
    # ================================================================

    def action_connect_dynamics(self):
        self.ensure_one()
        if not self.dynamics_org or not self.dynamics_region:
            raise UserError(_('Fill in Organization Name and Region first.'))

        # Validate org name — must be a simple identifier, no spaces
        org = self.dynamics_org.strip()
        if ' ' in org or '/' in org:
            raise UserError(_(
                'Organization Name must be your Dynamics 365 org identifier '
                '(e.g. "contoso"), not a display name.\n\n'
                'You can find it in your Dynamics 365 URL: '
                'https://<org-name>.crm.dynamics.com'
            ))

        scope = (
            f'https://{org}.api.'
            f'{self.dynamics_region}.dynamics.com/.default'
        )
        url = GRAPH_AUTH_URL.format(tenant=self.tenant_id) + '/token'
        data = {
            'client_id': self.client_id,
            'client_secret': self.client_secret,
            'grant_type': 'client_credentials',
            'scope': scope,
        }
        resp = requests.post(url, data=data, timeout=30)
        if resp.status_code != 200:
            raise UserError(_('Dynamics 365 authentication failed: %s') % resp.text)

        tokens = resp.json()
        self.write({
            'dynamics_access_token': tokens['access_token'],
            'dynamics_token_expiry': (
                fields.Datetime.now()
                + timedelta(seconds=tokens.get('expires_in', 3600))
            ),
            'state': (
                'fully_connected'
                if self.outlook_access_token
                else self.state
            ),
        })
        self._create_sync_log('dynamics_auth', 'success', 'Dynamics 365 CRM connected')

    def _ensure_dynamics_token(self):
        self.ensure_one()
        if not self.dynamics_access_token:
            raise UserError(_('Not connected to Dynamics 365. Click "Connect Dynamics".'))
        if self.dynamics_token_expiry and self.dynamics_token_expiry <= fields.Datetime.now():
            self.action_connect_dynamics()

    def _call_dynamics_api(self, method, endpoint, data=None, params=None):
        self._ensure_dynamics_token()
        headers = {
            'Authorization': f'Bearer {self.dynamics_access_token}',
            'Content-Type': 'application/json',
            'OData-MaxVersion': '4.0',
            'OData-Version': '4.0',
            'Prefer': 'odata.include-annotations="*"',
        }
        base = (
            f'https://{self.dynamics_org}.api.'
            f'{self.dynamics_region}.dynamics.com/api/data/v9.2'
        )
        url = f'{base}{endpoint}'
        try:
            resp = getattr(requests, method)(
                url, headers=headers, json=data, params=params, timeout=30
            )
            resp.raise_for_status()
            return resp.json() if resp.content else {}
        except requests.exceptions.HTTPError:
            _logger.error('Dynamics API %s %s → %s', method.upper(), endpoint, resp.text)
            raise UserError(_('Dynamics 365 API error: %s') % resp.text)
        except requests.exceptions.ConnectionError as exc:
            raise UserError(_('Cannot reach Dynamics 365: %s') % exc)

    # ================================================================
    #  TEST CONNECTION
    # ================================================================

    def action_test_connection(self):
        self.ensure_one()
        messages = []
        if self.outlook_access_token:
            try:
                info = self._call_graph_api('get', '/me')
                messages.append(
                    _('Outlook OK — connected as %s')
                    % info.get('mail', info.get('userPrincipalName', '?'))
                )
            except Exception as exc:
                messages.append(_('Outlook FAILED — %s') % exc)
        if self.dynamics_access_token:
            try:
                self._call_dynamics_api('get', '/WhoAmI')
                messages.append(_('Dynamics 365 CRM OK'))
            except Exception as exc:
                messages.append(_('Dynamics 365 FAILED — %s') % exc)
        if not messages:
            messages.append(_('No connections configured yet.'))
        return self._notify('\n'.join(messages))

    # ================================================================
    #  OUTLOOK — CONTACT SYNC (2-way)
    # ================================================================

    def action_sync_contacts(self):
        self.ensure_one()
        imported = exported = 0
        errors = []
        Partner = self.env['res.partner']

        # ── Import from Outlook → Odoo ──
        try:
            result = self._call_graph_api('get', '/me/contacts', params={
                '$top': 500,
                '$select': (
                    'id,displayName,givenName,surname,emailAddresses,'
                    'mobilePhone,businessPhones,companyName,jobTitle,'
                    'businessAddress'
                ),
            })
            for contact in result.get('value', []):
                ms_id = contact.get('id')
                vals = self._map_outlook_contact_to_partner(contact)
                existing = Partner.search([('microsoft365_id', '=', ms_id)], limit=1)
                if existing:
                    existing.write(vals)
                else:
                    vals['microsoft365_id'] = ms_id
                    Partner.create(vals)
                imported += 1
        except Exception as exc:
            errors.append(f'Import: {exc}')

        # ── Export from Odoo → Outlook ──
        partners = Partner.search([
            ('microsoft365_id', '=', False),
            ('email', '!=', False),
            ('is_company', '=', False),
        ], limit=200)
        for partner in partners:
            try:
                payload = self._map_partner_to_outlook_contact(partner)
                result = self._call_graph_api('post', '/me/contacts', data=payload)
                if result.get('id'):
                    partner.microsoft365_id = result['id']
                    exported += 1
            except Exception as exc:
                errors.append(f'Export {partner.name}: {exc}')

        details = f'Imported {imported}, exported {exported} contacts'
        if errors:
            details += '\n' + '\n'.join(errors)
        self._create_sync_log('contacts', 'error' if errors else 'success', details)
        self.last_sync = fields.Datetime.now()
        return self._notify(details)

    @staticmethod
    def _map_outlook_contact_to_partner(contact):
        emails = contact.get('emailAddresses') or []
        phones = contact.get('businessPhones') or []
        addr = contact.get('businessAddress') or {}
        return {
            'name': (
                contact.get('displayName')
                or f"{contact.get('givenName', '')} {contact.get('surname', '')}".strip()
                or 'Unknown'
            ),
            'email': emails[0].get('address') if emails else False,
            'phone': contact.get('mobilePhone') or (phones[0] if phones else False),
            'function': contact.get('jobTitle') or False,
            'company_name': contact.get('companyName') or False,
            'street': addr.get('street') or False,
            'city': addr.get('city') or False,
            'zip': addr.get('postalCode') or False,
            'is_company': False,
        }

    @staticmethod
    def _map_partner_to_outlook_contact(partner):
        data = {'displayName': partner.name}
        if partner.email:
            data['emailAddresses'] = [
                {'address': partner.email, 'name': partner.name}
            ]
        if partner.phone:
            data['businessPhones'] = [partner.phone]
        if partner.function:
            data['jobTitle'] = partner.function
        addr = {}
        if partner.street:
            addr['street'] = partner.street
        if partner.city:
            addr['city'] = partner.city
        if partner.zip:
            addr['postalCode'] = partner.zip
        if addr:
            data['businessAddress'] = addr
        return data

    # ================================================================
    #  OUTLOOK — CALENDAR SYNC (2-way)
    # ================================================================

    def action_sync_calendar(self):
        self.ensure_one()
        imported = exported = 0
        errors = []
        Event = self.env['calendar.event']

        # ── Import from Outlook ──
        try:
            result = self._call_graph_api('get', '/me/events', params={
                '$top': 200,
                '$orderby': 'start/dateTime desc',
                '$select': 'id,subject,bodyPreview,start,end,isAllDay,location',
            })
            for event in result.get('value', []):
                ms_id = event.get('id')
                vals = self._map_outlook_event_to_calendar(event)
                existing = Event.search([('microsoft365_id', '=', ms_id)], limit=1)
                if existing:
                    existing.write(vals)
                else:
                    vals['microsoft365_id'] = ms_id
                    Event.create(vals)
                imported += 1
        except Exception as exc:
            errors.append(f'Import: {exc}')

        # ── Export from Odoo ──
        events = Event.search([('microsoft365_id', '=', False)], limit=100)
        for ev in events:
            try:
                payload = self._map_calendar_to_outlook_event(ev)
                result = self._call_graph_api('post', '/me/events', data=payload)
                if result.get('id'):
                    ev.microsoft365_id = result['id']
                    exported += 1
            except Exception as exc:
                errors.append(f'Export {ev.name}: {exc}')

        details = f'Imported {imported}, exported {exported} events'
        if errors:
            details += '\n' + '\n'.join(errors)
        self._create_sync_log('calendar', 'error' if errors else 'success', details)
        self.last_sync = fields.Datetime.now()
        return self._notify(details)

    @staticmethod
    def _parse_graph_datetime(dt_str):
        """Parse datetime strings from Microsoft Graph (may have 7 decimal places)."""
        if not dt_str:
            return None
        dt_str = dt_str.rstrip('Z')
        # Python fromisoformat supports up to 6 decimals; Graph sends 7
        if '.' in dt_str:
            base, frac = dt_str.rsplit('.', 1)
            dt_str = f'{base}.{frac[:6]}'
        return datetime.fromisoformat(dt_str)

    @staticmethod
    def _map_outlook_event_to_calendar(event):
        vals = {
            'name': event.get('subject') or 'No Subject',
            'description': event.get('bodyPreview') or False,
            'allday': event.get('isAllDay', False),
        }
        start = event.get('start') or {}
        end = event.get('end') or {}
        if start.get('dateTime'):
            vals['start'] = Microsoft365Connection._parse_graph_datetime(start['dateTime'])
        if end.get('dateTime'):
            vals['stop'] = Microsoft365Connection._parse_graph_datetime(end['dateTime'])
        loc = event.get('location') or {}
        if loc.get('displayName'):
            vals['location'] = loc['displayName']
        return vals

    @staticmethod
    def _map_calendar_to_outlook_event(ev):
        data = {
            'subject': ev.name or 'Odoo Event',
            'body': {'contentType': 'text', 'content': ev.description or ''},
            'isAllDay': ev.allday,
        }
        if ev.start:
            start_dt = ev.start
            data['start'] = {
                'dateTime': start_dt.isoformat(),
                'timeZone': 'UTC',
            }
            if ev.allday:
                # All-day events need at least 24h duration
                end_dt = ev.stop if ev.stop and ev.stop > start_dt else (start_dt + timedelta(days=1))
                if (end_dt - start_dt).total_seconds() < 86400:
                    end_dt = start_dt + timedelta(days=1)
                data['end'] = {
                    'dateTime': end_dt.isoformat(),
                    'timeZone': 'UTC',
                }
            elif ev.stop:
                data['end'] = {
                    'dateTime': ev.stop.isoformat(),
                    'timeZone': 'UTC',
                }
        if ev.location:
            data['location'] = {'displayName': ev.location}
        return data

    # ================================================================
    #  OUTLOOK — EMAIL IMPORT
    # ================================================================

    def action_import_emails(self):
        self.ensure_one()
        imported = 0
        rules_fired = 0
        errors = []
        result = {}
        MailMsg = self.env['mail.message']
        Attachment = self.env['ir.attachment']

        try:
            result = self._call_graph_api('get', '/me/messages', params={
                '$top': 50,
                '$orderby': 'receivedDateTime desc',
                '$select': (
                    'id,subject,bodyPreview,body,from,toRecipients,'
                    'receivedDateTime,isRead,hasAttachments,'
                    'conversationId'
                ),
            })
            for msg in result.get('value', []):
                ms_id = msg.get('id')
                if MailMsg.search_count([('microsoft365_id', '=', ms_id)]):
                    continue

                from_info = (msg.get('from') or {}).get('emailAddress') or {}
                from_email = from_info.get('address', '')
                partner = self.env['res.partner'].search(
                    [('email', '=ilike', from_email)], limit=1
                )

                # Full HTML body if available, fall back to preview
                body_obj = msg.get('body') or {}
                body_html = body_obj.get('content') or msg.get('bodyPreview', '')

                # Parse ISO datetime from Microsoft
                raw_date = msg.get('receivedDateTime')
                if raw_date:
                    try:
                        parsed_date = self._parse_graph_datetime(
                            raw_date.replace('Z', '')
                        ).strftime('%Y-%m-%d %H:%M:%S')
                    except (ValueError, AttributeError):
                        parsed_date = fields.Datetime.now()
                else:
                    parsed_date = fields.Datetime.now()

                new_msg = MailMsg.sudo().create({
                    'subject': msg.get('subject', ''),
                    'body': body_html,
                    'email_from': from_email,
                    'date': parsed_date,
                    'message_type': 'email',
                    'subtype_id': self.env.ref('mail.mt_note').id,
                    'model': 'res.partner' if partner else 'sm.365.connection',
                    'res_id': partner.id if partner else self.id,
                    'author_id': partner.id if partner else False,
                    'microsoft365_id': ms_id,
                    'microsoft365_conversation_id': msg.get('conversationId'),
                    'microsoft365_is_read': msg.get('isRead', False),
                    'microsoft365_has_attachments': msg.get('hasAttachments', False),
                })

                # Download attachments — check explicit flag OR cid: in body
                has_inline = 'cid:' in body_html
                if msg.get('hasAttachments') or has_inline:
                    try:
                        att_result = self._call_graph_api(
                            'get', f'/me/messages/{ms_id}/attachments',
                            params={'$top': 100},
                        )
                        cid_map = {}  # contentId -> Odoo attachment URL
                        for att in att_result.get('value', []):
                            if att.get('@odata.type') != '#microsoft.graph.fileAttachment':
                                continue
                            content = att.get('contentBytes')
                            if not content:
                                continue
                            att_rec = Attachment.sudo().create({
                                'name': att.get('name') or 'attachment',
                                'datas': content,
                                'res_model': 'mail.message',
                                'res_id': new_msg.id,
                                'mimetype': att.get('contentType') or 'application/octet-stream',
                            })
                            # Map inline images (cid:xxx) to Odoo URLs
                            content_id = att.get('contentId')
                            if content_id and att.get('isInline'):
                                att_url = f'/web/image/{att_rec.id}'
                                cid_map[content_id] = att_url

                        # Replace cid: references in body with Odoo attachment URLs
                        if cid_map and new_msg.body:
                            updated_body = new_msg.body
                            for cid, att_url in cid_map.items():
                                updated_body = updated_body.replace(
                                    f'cid:{cid}', att_url
                                )
                            if updated_body != new_msg.body:
                                new_msg.sudo().write({'body': updated_body})
                    except Exception as exc:
                        _logger.warning('Attachment fetch failed for %s: %s', ms_id, exc)

                # Apply email rules
                for rule in self.email_rule_ids.filtered('active'):
                    if rule._match(msg):
                        try:
                            rule._apply(msg, new_msg)
                            rules_fired += 1
                            break
                        except Exception as exc:
                            _logger.warning('Rule %s failed: %s', rule.name, exc)

                imported += 1
        except Exception as exc:
            errors.append(str(exc))

        skipped = 0
        try:
            total_fetched = len(result.get('value', []))
            skipped = total_fetched - imported
        except Exception:
            pass
        details = f'Imported {imported} new emails'
        if skipped:
            details += f' ({skipped} already synced)'
        if rules_fired:
            details += f', {rules_fired} rules fired'
        if errors:
            details += '\n' + '\n'.join(errors)
        self._create_sync_log('emails', 'error' if errors else 'success', details)
        self.last_sync = fields.Datetime.now()
        return self._notify(details)

    def action_refetch_email_images(self):
        """Re-download inline images for existing emails that have cid: references."""
        self.ensure_one()
        MailMsg = self.env['mail.message']
        Attachment = self.env['ir.attachment']
        fixed = 0

        msgs = MailMsg.sudo().search([
            ('microsoft365_id', '!=', False),
            ('body', 'like', 'cid:'),
        ])
        for msg_rec in msgs:
            ms_id = msg_rec.microsoft365_id
            try:
                att_result = self._call_graph_api(
                    'get', f'/me/messages/{ms_id}/attachments',
                    params={'$top': 100},
                )
                cid_map = {}
                for att in att_result.get('value', []):
                    if att.get('@odata.type') != '#microsoft.graph.fileAttachment':
                        continue
                    content = att.get('contentBytes')
                    content_id = att.get('contentId')
                    if not content or not content_id:
                        continue
                    if not att.get('isInline'):
                        continue
                    # Check if we already have this attachment
                    existing = Attachment.sudo().search([
                        ('res_model', '=', 'mail.message'),
                        ('res_id', '=', msg_rec.id),
                        ('name', '=', att.get('name', content_id)),
                    ], limit=1)
                    if existing:
                        att_url = f'/web/image/{existing.id}'
                    else:
                        att_rec = Attachment.sudo().create({
                            'name': att.get('name') or content_id,
                            'datas': content,
                            'res_model': 'mail.message',
                            'res_id': msg_rec.id,
                            'mimetype': att.get('contentType') or 'image/png',
                        })
                        att_url = f'/web/image/{att_rec.id}'
                    cid_map[content_id] = att_url

                if cid_map:
                    updated_body = msg_rec.body
                    for cid, att_url in cid_map.items():
                        updated_body = updated_body.replace(f'cid:{cid}', att_url)
                    if updated_body != msg_rec.body:
                        msg_rec.sudo().write({'body': updated_body})
                        fixed += 1
            except Exception as exc:
                _logger.warning('Refetch images for %s failed: %s', ms_id, exc)

        return self._notify(_('Fixed inline images for %d emails') % fixed)

    # ================================================================
    #  OUTLOOK — SEND EMAIL
    # ================================================================

    def send_email(self, to_email, subject, body_html):
        """Send an email via Microsoft Graph API (Outlook)."""
        self.ensure_one()
        payload = {
            'message': {
                'subject': subject,
                'body': {'contentType': 'HTML', 'content': body_html},
                'toRecipients': [
                    {'emailAddress': {'address': to_email}}
                ],
            },
            'saveToSentItems': 'true',
        }
        self._call_graph_api('post', '/me/sendMail', data=payload)
        self._create_sync_log(
            'send_email', 'success', f'Sent to {to_email}: {subject}'
        )

    # ================================================================
    #  DYNAMICS 365 — IMPORT LEADS
    # ================================================================

    def action_import_crm_leads(self):
        self.ensure_one()
        imported = updated = 0
        errors = []
        Lead = self.env['crm.lead']

        try:
            result = self._call_dynamics_api('get', '/leads', params={
                '$select': (
                    'leadid,subject,firstname,lastname,emailaddress1,'
                    'telephone1,companyname,description,revenue'
                ),
                '$top': 500,
            })
            for lead in result.get('value', []):
                d_id = lead.get('leadid')
                contact = (
                    f"{lead.get('firstname', '')} "
                    f"{lead.get('lastname', '')}".strip() or False
                )
                vals = {
                    'name': lead.get('subject') or 'Imported Lead',
                    'contact_name': contact,
                    'email_from': lead.get('emailaddress1'),
                    'phone': lead.get('telephone1'),
                    'partner_name': lead.get('companyname'),
                    'description': lead.get('description'),
                    'expected_revenue': lead.get('revenue', 0),
                    'type': 'lead',
                }
                existing = Lead.search([('dynamics365_id', '=', d_id)], limit=1)
                if existing:
                    existing.write(vals)
                    updated += 1
                else:
                    vals['dynamics365_id'] = d_id
                    Lead.create(vals)
                    imported += 1
        except Exception as exc:
            errors.append(str(exc))

        details = f'Imported {imported}, updated {updated} leads'
        if errors:
            details += '\n' + '\n'.join(errors)
        self._create_sync_log('crm_leads', 'error' if errors else 'success', details)
        return self._notify(details)

    # ================================================================
    #  DYNAMICS 365 — IMPORT OPPORTUNITIES
    # ================================================================

    def action_import_crm_opportunities(self):
        self.ensure_one()
        imported = updated = 0
        errors = []
        Lead = self.env['crm.lead']

        try:
            result = self._call_dynamics_api('get', '/opportunities', params={
                '$select': (
                    'opportunityid,name,estimatedvalue,'
                    'description,closeprobability'
                ),
                '$top': 500,
            })
            for opp in result.get('value', []):
                d_id = opp.get('opportunityid')
                vals = {
                    'name': opp.get('name') or 'Imported Opportunity',
                    'expected_revenue': opp.get('estimatedvalue', 0),
                    'description': opp.get('description'),
                    'probability': opp.get('closeprobability', 0),
                    'type': 'opportunity',
                }
                existing = Lead.search([('dynamics365_id', '=', d_id)], limit=1)
                if existing:
                    existing.write(vals)
                    updated += 1
                else:
                    vals['dynamics365_id'] = d_id
                    Lead.create(vals)
                    imported += 1
        except Exception as exc:
            errors.append(str(exc))

        details = f'Imported {imported}, updated {updated} opportunities'
        if errors:
            details += '\n' + '\n'.join(errors)
        self._create_sync_log(
            'crm_opportunities', 'error' if errors else 'success', details
        )
        return self._notify(details)

    # ================================================================
    #  DYNAMICS 365 — IMPORT ACCOUNTS
    # ================================================================

    def action_import_crm_accounts(self):
        self.ensure_one()
        imported = updated = 0
        errors = []
        Partner = self.env['res.partner']

        try:
            result = self._call_dynamics_api('get', '/accounts', params={
                '$select': (
                    'accountid,name,emailaddress1,telephone1,'
                    'address1_line1,address1_city,address1_postalcode,'
                    'websiteurl'
                ),
                '$top': 500,
            })
            for acct in result.get('value', []):
                d_id = acct.get('accountid')
                vals = {
                    'name': acct.get('name') or 'Imported Account',
                    'is_company': True,
                    'email': acct.get('emailaddress1'),
                    'phone': acct.get('telephone1'),
                    'street': acct.get('address1_line1'),
                    'city': acct.get('address1_city'),
                    'zip': acct.get('address1_postalcode'),
                    'website': acct.get('websiteurl'),
                }
                existing = Partner.search([('dynamics365_id', '=', d_id)], limit=1)
                if existing:
                    existing.write(vals)
                    updated += 1
                else:
                    vals['dynamics365_id'] = d_id
                    Partner.create(vals)
                    imported += 1
        except Exception as exc:
            errors.append(str(exc))

        details = f'Imported {imported}, updated {updated} accounts'
        if errors:
            details += '\n' + '\n'.join(errors)
        self._create_sync_log(
            'crm_accounts', 'error' if errors else 'success', details
        )
        return self._notify(details)

    # ================================================================
    #  DYNAMICS 365 — IMPORT CONTACTS (individual people)
    # ================================================================

    def action_import_crm_contacts(self):
        self.ensure_one()
        imported = updated = 0
        errors = []
        Partner = self.env['res.partner']

        try:
            result = self._call_dynamics_api('get', '/contacts', params={
                '$select': (
                    'contactid,fullname,firstname,lastname,emailaddress1,'
                    'telephone1,mobilephone,jobtitle,address1_line1,'
                    'address1_city,address1_postalcode'
                ),
                '$top': 500,
            })
            for c in result.get('value', []):
                d_id = c.get('contactid')
                name = (
                    c.get('fullname')
                    or f"{c.get('firstname', '')} {c.get('lastname', '')}".strip()
                    or 'Imported Contact'
                )
                vals = {
                    'name': name,
                    'is_company': False,
                    'email': c.get('emailaddress1'),
                    'phone': c.get('telephone1'),
                    # mobilephone mapped to phone if phone is empty
                    'function': c.get('jobtitle'),
                    'street': c.get('address1_line1'),
                    'city': c.get('address1_city'),
                    'zip': c.get('address1_postalcode'),
                }
                existing = Partner.search(
                    [('dynamics365_id', '=', d_id)], limit=1
                )
                if existing:
                    existing.write(vals)
                    updated += 1
                else:
                    vals['dynamics365_id'] = d_id
                    Partner.create(vals)
                    imported += 1
        except Exception as exc:
            errors.append(str(exc))

        details = f'Imported {imported}, updated {updated} CRM contacts'
        if errors:
            details += '\n' + '\n'.join(errors)
        self._create_sync_log(
            'crm_contacts', 'error' if errors else 'success', details
        )
        return self._notify(details)

    # ================================================================
    #  SCHEDULED AUTO-SYNC (Cron)
    # ================================================================

    @api.model
    def _cron_sync_all(self):
        for conn in self.search([('auto_sync', '=', True)]):
            try:
                if conn.outlook_access_token:
                    if conn.sync_contacts:
                        conn.action_sync_contacts()
                    if conn.sync_calendar:
                        conn.action_sync_calendar()
                    if conn.sync_emails:
                        conn.action_import_emails()
                if (
                    conn.dynamics_enabled
                    and conn.dynamics_access_token
                    and conn.sync_crm
                ):
                    conn.action_import_crm_leads()
                    conn.action_import_crm_opportunities()
                    conn.action_import_crm_accounts()
                    conn.action_import_crm_contacts()
                conn.last_sync = fields.Datetime.now()
            except Exception as exc:
                _logger.error('Auto-sync %s failed: %s', conn.name, exc)
                conn._create_sync_log('auto_sync', 'error', str(exc))

    # ================================================================
    #  UTILITIES
    # ================================================================

    def _create_sync_log(self, sync_type, status, details):
        self.env['sm.365.sync.log'].sudo().create({
            'connection_id': self.id,
            'sync_type': sync_type,
            'status': status,
            'details': details,
        })

    def _notify(self, message):
        return {
            'type': 'ir.actions.client',
            'tag': 'display_notification',
            'params': {
                'title': _('Microsoft 365'),
                'message': message,
                'type': 'success',
                'sticky': False,
            },
        }

    def action_view_sync_logs(self):
        return {
            'type': 'ir.actions.act_window',
            'name': _('Sync Logs'),
            'res_model': 'sm.365.sync.log',
            'view_mode': 'list,form',
            'domain': [('connection_id', '=', self.id)],
        }

    def action_disconnect(self):
        self.write({
            'outlook_access_token': False,
            'outlook_refresh_token': False,
            'outlook_token_expiry': False,
            'outlook_user_email': False,
            'dynamics_access_token': False,
            'dynamics_token_expiry': False,
            'state': 'draft',
        })
        # Disable SSO if no other active connections remain
        other_active = self.search([
            ('id', '!=', self.id),
            ('state', 'in', ('outlook_connected', 'fully_connected')),
        ], limit=1)
        if not other_active:
            self._disable_sso_provider()

    # ================================================================
    #  SSO PROVIDER AUTO-CONFIGURATION
    # ================================================================

    def _get_sso_provider(self):
        """Return the Microsoft 365 SSO provider record created by this module."""
        provider = self.env.ref(
            'sm_365_connector.provider_microsoft_365', raise_if_not_found=False,
        )
        if not provider:
            provider = self.env['auth.oauth.provider'].sudo().search(
                [('name', '=', 'Microsoft 365')], limit=1,
            )
        return provider

    def _configure_sso_provider(self):
        """Enable the Microsoft 365 SSO provider and set its Client ID and
        tenant-specific endpoints from this connection's credentials."""
        self.ensure_one()
        provider = self._get_sso_provider()
        if not provider:
            return

        tenant = self.tenant_id or 'common'
        provider.sudo().write({
            'client_id': self.client_id,
            'auth_endpoint': (
                f'https://login.microsoftonline.com/{tenant}'
                '/oauth2/v2.0/authorize'
            ),
            'enabled': True,
        })
        _logger.info(
            'SSO provider "%s" auto-configured with Client ID %s (tenant: %s)',
            provider.name, self.client_id, tenant,
        )

    def _disable_sso_provider(self):
        """Disable the Microsoft 365 SSO provider."""
        provider = self._get_sso_provider()
        if provider and provider.enabled:
            provider.sudo().write({'enabled': False})
            _logger.info('SSO provider "%s" disabled', provider.name)

    # ================================================================
    #  TEAMS MEETINGS
    # ================================================================

    def _create_teams_meeting_for_event(self, event):
        """Create an onlineMeeting in Teams and link it to a calendar.event."""
        self.ensure_one()
        start = event.start.isoformat() if event.start else fields.Datetime.now().isoformat()
        end = event.stop.isoformat() if event.stop else (
            fields.Datetime.now() + timedelta(hours=1)
        ).isoformat()
        payload = {
            'startDateTime': start + 'Z' if not start.endswith('Z') else start,
            'endDateTime': end + 'Z' if not end.endswith('Z') else end,
            'subject': event.name or 'Odoo Meeting',
        }
        try:
            result = self._call_graph_api(
                'post', '/me/onlineMeetings', data=payload,
            )
        except UserError as exc:
            err = str(exc)
            if 'AuthenticationError' in err or 'authenticating with resource' in err.lower():
                raise UserError(_(
                    'Teams Meetings require a Microsoft 365 Business/Enterprise '
                    'license with Teams enabled.\n\n'
                    'Personal Outlook.com / Hotmail accounts do not have access '
                    'to the Teams Online Meetings API.\n\n'
                    'To use this feature, connect with a work/school M365 account.'
                ))
            raise
        event.sudo().write({
            'is_teams_meeting': True,
            'teams_join_url': result.get('joinUrl'),
            'teams_conference_id': result.get('id'),
        })
        self._create_sync_log(
            'calendar', 'success',
            f'Created Teams meeting for event "{event.name}"',
        )
        return result

    # ================================================================
    #  MICROSOFT TO DO  (Tasks.ReadWrite)
    # ================================================================

    def _get_default_todo_list_id(self):
        """Fetch the default Microsoft To Do list (Tasks)."""
        self.ensure_one()
        result = self._call_graph_api('get', '/me/todo/lists')
        for lst in result.get('value', []):
            if lst.get('wellknownListName') == 'defaultList':
                return lst.get('id')
        if result.get('value'):
            return result['value'][0].get('id')
        return False

    def action_sync_todo_tasks(self):
        """2-way sync project.task ↔ Microsoft To Do."""
        self.ensure_one()
        imported = exported = 0
        errors = []
        Task = self.env['project.task']
        list_id = self._get_default_todo_list_id()
        if not list_id:
            return self._notify(_('No To Do list found on Microsoft account.'))

        # ── Import from To Do → Odoo ──
        try:
            result = self._call_graph_api(
                'get', f'/me/todo/lists/{list_id}/tasks',
                params={'$top': 200},
            )
            for t in result.get('value', []):
                ms_id = t.get('id')
                vals = {
                    'name': t.get('title') or 'Untitled Task',
                    'description': (t.get('body') or {}).get('content') or '',
                    'microsoft365_list_id': list_id,
                }
                if t.get('dueDateTime'):
                    try:
                        dt = Microsoft365Connection._parse_graph_datetime(
                            t['dueDateTime']['dateTime']
                        )
                        # date_deadline is fields.Datetime in Odoo 19
                        vals['date_deadline'] = dt.strftime('%Y-%m-%d %H:%M:%S')
                    except Exception:
                        pass
                # Map status
                if t.get('status') == 'completed':
                    vals['state'] = '1_done'

                existing = Task.search([('microsoft365_id', '=', ms_id)], limit=1)
                if existing:
                    existing.write(vals)
                else:
                    vals['microsoft365_id'] = ms_id
                    Task.create(vals)
                imported += 1
        except Exception as exc:
            errors.append(f'Import: {exc}')

        # ── Export from Odoo → To Do ──
        tasks = Task.search([
            ('microsoft365_id', '=', False),
            ('active', '=', True),
        ], limit=100)
        for task in tasks:
            try:
                payload = {
                    'title': task.name or 'Untitled',
                    'body': {
                        'content': task.description or '',
                        'contentType': 'text',
                    },
                }
                if task.date_deadline:
                    # date_deadline is datetime in Odoo 19; Graph expects YYYY-MM-DDT00:00:00.0000000
                    dl = task.date_deadline
                    dl_str = dl.strftime('%Y-%m-%dT00:00:00.0000000') if dl else None
                    if dl_str:
                        payload['dueDateTime'] = {
                            'dateTime': dl_str,
                            'timeZone': 'UTC',
                        }
                result = self._call_graph_api(
                    'post', f'/me/todo/lists/{list_id}/tasks', data=payload,
                )
                if result.get('id'):
                    task.sudo().write({
                        'microsoft365_id': result['id'],
                        'microsoft365_list_id': list_id,
                    })
                    exported += 1
            except Exception as exc:
                errors.append(f'Export {task.name}: {exc}')

        details = f'Imported {imported}, exported {exported} tasks'
        if errors:
            details += '\n' + '\n'.join(errors[:5])
        self._create_sync_log(
            'tasks',
            'error' if errors else 'success', details,
        )
        self.last_sync = fields.Datetime.now()
        return self._notify(details)

    # ================================================================
    #  ONEDRIVE ATTACHMENT BACKEND  (Files.ReadWrite)
    # ================================================================

    def _onedrive_ensure_folder(self, folder_path):
        """Ensure a folder exists in OneDrive. Returns the folder item dict."""
        self.ensure_one()
        path = (folder_path or '/').strip('/')
        if not path:
            return self._call_graph_api('get', '/me/drive/root')
        try:
            return self._call_graph_api('get', f'/me/drive/root:/{path}')
        except Exception:
            # Create folder(s) recursively
            parts = path.split('/')
            parent = self._call_graph_api('get', '/me/drive/root')
            cur_path = ''
            for part in parts:
                cur_path = f'{cur_path}/{part}' if cur_path else part
                try:
                    parent = self._call_graph_api('get', f'/me/drive/root:/{cur_path}')
                except Exception:
                    parent = self._call_graph_api(
                        'post', f'/me/drive/items/{parent["id"]}/children',
                        data={
                            'name': part,
                            'folder': {},
                            '@microsoft.graph.conflictBehavior': 'replace',
                        },
                    )
            return parent

    @staticmethod
    def _sanitize_onedrive_filename(name):
        """Remove characters not allowed in OneDrive file/folder names."""
        import re as _re
        # OneDrive forbidden: " * : < > ? / \ |
        sanitized = _re.sub(r'[\"*:<>?/\\|]', '_', name)
        # Collapse multiple underscores / spaces
        sanitized = _re.sub(r'_{2,}', '_', sanitized)
        sanitized = sanitized.strip(' _.')
        return sanitized[:250] or f'file_{id(name)}'

    def _onedrive_upload_attachment(self, attachment):
        """Upload an ir.attachment to OneDrive and store back the item_id/URLs."""
        self.ensure_one()
        self._ensure_outlook_token()
        folder_path = (self.onedrive_folder_path or '/Apps/Odoo').strip('/')
        raw_name = attachment.name or f'attachment_{attachment.id}'
        filename = self._sanitize_onedrive_filename(raw_name)
        upload_path = f'/me/drive/root:/{folder_path}/{filename}:/content'
        headers = {
            'Authorization': f'Bearer {self.outlook_access_token}',
            'Content-Type': attachment.mimetype or 'application/octet-stream',
        }
        url = f'{GRAPH_API_URL}{upload_path}'
        resp = requests.put(
            url, headers=headers, data=base64.b64decode(attachment.datas or b''),
            timeout=120,
        )
        if resp.status_code not in (200, 201):
            raise UserError(_('OneDrive upload failed: %s') % resp.text)
        item = resp.json()
        attachment.sudo().write({
            'onedrive_item_id': item.get('id'),
            'onedrive_web_url': item.get('webUrl'),
            'onedrive_download_url': item.get(
                '@microsoft.graph.downloadUrl'
            ) or item.get('webUrl'),
        })
        return item

    def action_upload_attachments_to_onedrive(self):
        """Upload all local attachments that aren't yet on OneDrive."""
        self.ensure_one()
        if not self.onedrive_enabled:
            raise UserError(_('Enable OneDrive in the configuration tab first.'))
        Attachment = self.env['ir.attachment']
        # Only upload attachments linked to real business records
        BUSINESS_MODELS = (
            'sale.order', 'purchase.order', 'account.move',
            'stock.picking', 'project.project', 'project.task',
            'hr.expense', 'crm.lead', 'helpdesk.ticket',
            'product.template', 'product.product',
            'res.partner', 'mail.compose.message',
        )
        domain = [
            ('onedrive_item_id', '=', False),
            ('store_fname', '!=', False),
            ('file_size', '>', 0),
            ('file_size', '<', 4 * 1024 * 1024),  # max 4MB per simple upload
            ('res_model', 'in', BUSINESS_MODELS),
        ]
        atts = Attachment.search(domain, limit=20)
        if not atts:
            return self._notify(_(
                'No pending business attachments to upload (< 4MB).\n'
                'Only attachments from Sales, Purchases, Invoices, '
                'Inventory, Projects, CRM, etc. are synced.'
            ))
        uploaded = 0
        errors = []
        for att in atts:
            try:
                self._onedrive_upload_attachment(att)
                uploaded += 1
            except Exception as exc:
                errors.append(f'{att.name}: {exc}')
                if len(errors) >= 3:
                    break  # stop after 3 errors to avoid long freeze
        remaining = Attachment.search_count(domain)
        details = f'Uploaded {uploaded} attachments to OneDrive'
        if remaining:
            details += f' ({remaining} still pending — click again)'
        if errors:
            details += '\n' + '\n'.join(errors[:5])
        self._create_sync_log(
            'send_email', 'error' if errors else 'success', details,
        )
        return self._notify(details)

    # ================================================================
    #  GRAPH WEBHOOK SUBSCRIPTIONS  (real-time push)
    # ================================================================

    def _get_webhook_url(self):
        base_url = self.env['ir.config_parameter'].sudo().get_param('web.base.url')
        return f'{base_url}/microsoft365/webhook'

    def _graph_subscribe(self, sub):
        """Create a Graph subscription for push notifications."""
        self.ensure_one()
        from datetime import datetime as _dt
        expiration = _dt.utcnow() + timedelta(days=2, hours=20)
        client_state = sub.client_state or sub._generate_client_state()
        payload = {
            'changeType': 'created,updated',
            'notificationUrl': self._get_webhook_url(),
            'resource': sub.resource,
            'expirationDateTime': expiration.strftime('%Y-%m-%dT%H:%M:%S.0000000Z'),
            'clientState': client_state,
        }
        result = self._call_graph_api('post', '/subscriptions', data=payload)
        sub.sudo().write({
            'subscription_id': result.get('id'),
            'client_state': client_state,
            'expiration': fields.Datetime.now() + timedelta(days=2, hours=20),
            'state': 'active',
        })

    def _graph_unsubscribe(self, sub):
        self.ensure_one()
        if sub.subscription_id:
            try:
                self._call_graph_api(
                    'delete', f'/subscriptions/{sub.subscription_id}',
                )
            except Exception as exc:
                _logger.warning('Unsubscribe failed: %s', exc)
        sub.sudo().unlink()

    def _graph_renew(self, sub):
        self.ensure_one()
        from datetime import datetime as _dt
        expiration = _dt.utcnow() + timedelta(days=2, hours=20)
        self._call_graph_api(
            'patch', f'/subscriptions/{sub.subscription_id}',
            data={
                'expirationDateTime': expiration.strftime(
                    '%Y-%m-%dT%H:%M:%S.0000000Z'
                ),
            },
        )
        sub.sudo().write({
            'expiration': fields.Datetime.now() + timedelta(days=2, hours=20),
            'state': 'active',
        })

    def action_enable_webhooks(self):
        """Create subscriptions for enabled sync types."""
        self.ensure_one()
        webhook_url = self._get_webhook_url()
        if not webhook_url.startswith('https://'):
            raise UserError(_(
                'Microsoft Graph webhooks require a public HTTPS URL.\n\n'
                'Your current web.base.url is: %s\n\n'
                'To fix this:\n'
                '1. Use a tunnel like ngrok: ngrok http 8019\n'
                '2. Update Settings → System Parameters → web.base.url to the https:// URL\n'
                '3. Then try enabling webhooks again.'
            ) % webhook_url)

        Sub = self.env['sm.365.subscription']
        # Clean up old error/draft subscriptions first
        old_broken = Sub.search([
            ('connection_id', '=', self.id),
            ('state', 'in', ('error', 'draft')),
        ])
        if old_broken:
            old_broken.sudo().unlink()

        resources = []
        if self.sync_emails:
            resources.append('me/messages')
        if self.sync_contacts:
            resources.append('me/contacts')
        if self.sync_calendar:
            resources.append('me/events')

        created = 0
        errors = []
        for res in resources:
            existing = Sub.search([
                ('connection_id', '=', self.id),
                ('resource', '=', res),
                ('state', '=', 'active'),
            ], limit=1)
            if existing:
                continue
            sub = Sub.create({
                'connection_id': self.id,
                'resource': res,
            })
            try:
                self._graph_subscribe(sub)
                created += 1
            except Exception as exc:
                errors.append(f'{res}: {exc}')
                sub.state = 'error'

        self.webhooks_enabled = True
        details = f'Created {created} Graph subscriptions'
        if errors:
            details += '\n' + '\n'.join(errors)
        self._create_sync_log(
            'outlook_auth', 'error' if errors else 'success', details,
        )
        return self._notify(details)

    def action_disable_webhooks(self):
        self.ensure_one()
        for sub in self.subscription_ids:
            self._graph_unsubscribe(sub)
        self.webhooks_enabled = False
        return self._notify(_('Webhooks disabled.'))

    # ================================================================
    #  DASHBOARD / RETRY
    # ================================================================

    def action_retry_failed(self):
        """Re-run sync types that have failed recently."""
        self.ensure_one()
        Log = self.env['sm.365.sync.log']
        recent_errors = Log.search([
            ('connection_id', '=', self.id),
            ('status', '=', 'error'),
        ], limit=20, order='id desc')
        sync_types = set(recent_errors.mapped('sync_type'))
        dispatch = {
            'contacts': 'action_sync_contacts',
            'calendar': 'action_sync_calendar',
            'emails': 'action_import_emails',
            'tasks': 'action_sync_todo_tasks',
            'crm_leads': 'action_import_crm_leads',
            'crm_opportunities': 'action_import_crm_opportunities',
            'crm_accounts': 'action_import_crm_accounts',
            'crm_contacts': 'action_import_crm_contacts',
        }
        done = []
        for st in sync_types:
            fn = dispatch.get(st)
            if fn and hasattr(self, fn):
                try:
                    getattr(self, fn)()
                    done.append(st)
                except Exception as exc:
                    _logger.warning('Retry %s failed: %s', st, exc)
        # Mark old error logs as resolved for successful retries
        if done:
            resolved_logs = recent_errors.filtered(
                lambda l: l.sync_type in done
            )
            resolved_logs.sudo().unlink()
        return self._notify(
            _('Retried: %s') % (', '.join(done) if done else _('nothing'))
        )

    def action_view_email_rules(self):
        self.ensure_one()
        return {
            'type': 'ir.actions.act_window',
            'name': _('Email Rules'),
            'res_model': 'sm.365.email.rule',
            'view_mode': 'list,form',
            'domain': [('connection_id', '=', self.id)],
            'context': {'default_connection_id': self.id},
        }

    def action_view_subscriptions(self):
        self.ensure_one()
        return {
            'type': 'ir.actions.act_window',
            'name': _('Graph Subscriptions'),
            'res_model': 'sm.365.subscription',
            'view_mode': 'list,form',
            'domain': [('connection_id', '=', self.id)],
            'context': {'default_connection_id': self.id},
        }
