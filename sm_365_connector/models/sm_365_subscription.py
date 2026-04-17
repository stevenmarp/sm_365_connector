# -*- coding: utf-8 -*-

import logging
import secrets
from datetime import timedelta

from odoo import api, fields, models, _
from odoo.exceptions import UserError

_logger = logging.getLogger(__name__)


class Microsoft365Subscription(models.Model):
    """Microsoft Graph change notification subscription (webhook).

    Graph subscriptions require the Odoo server to be reachable via HTTPS
    from the public internet. The endpoint URL used is:
        {web.base.url}/microsoft365/webhook

    Subscriptions expire (max ~3 days for most resources) and must be renewed.
    A scheduled action renews them automatically.
    """

    _name = 'sm.365.subscription'
    _description = 'Microsoft Graph Subscription'
    _order = 'id desc'

    connection_id = fields.Many2one(
        'sm.365.connection', required=True, ondelete='cascade', index=True,
    )
    resource = fields.Selection([
        ('me/messages', 'Inbox (Emails)'),
        ('me/contacts', 'Contacts'),
        ('me/events', 'Calendar Events'),
    ], required=True)
    subscription_id = fields.Char('Graph Subscription ID', readonly=True, index=True)
    client_state = fields.Char(
        'Client State (secret)', readonly=True, copy=False,
        help='Used to validate incoming webhook notifications.',
    )
    expiration = fields.Datetime('Expires At', readonly=True)
    state = fields.Selection([
        ('draft', 'Draft'),
        ('active', 'Active'),
        ('expired', 'Expired'),
        ('error', 'Error'),
    ], default='draft', readonly=True)
    last_notification = fields.Datetime(readonly=True)
    notification_count = fields.Integer(readonly=True, default=0)

    def action_subscribe(self):
        for sub in self:
            sub.connection_id._graph_subscribe(sub)

    def action_unsubscribe(self):
        for sub in self:
            sub.connection_id._graph_unsubscribe(sub)

    def action_renew(self):
        for sub in self:
            sub.connection_id._graph_renew(sub)

    @api.model
    def _generate_client_state(self):
        return secrets.token_urlsafe(32)

    @api.model
    def _cron_renew(self):
        """Renew subscriptions nearing expiry (< 6 hours left)."""
        threshold = fields.Datetime.now() + timedelta(hours=6)
        expiring = self.search([
            ('state', '=', 'active'),
            ('expiration', '<=', threshold),
        ])
        for sub in expiring:
            try:
                sub.connection_id._graph_renew(sub)
            except Exception as exc:
                _logger.error('Failed to renew subscription %s: %s', sub.id, exc)
                sub.state = 'error'
