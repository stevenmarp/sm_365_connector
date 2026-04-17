# -*- coding: utf-8 -*-

from odoo import api, fields, models, _
from odoo.exceptions import UserError


class OutlookSendEmail(models.TransientModel):
    _name = 'sm.365.send.email'
    _description = 'Send Email via Outlook'

    connection_id = fields.Many2one(
        'sm.365.connection', string='Connection', required=True,
        default=lambda self: self._default_connection(),
    )
    partner_id = fields.Many2one('res.partner', string='Recipient')
    to_email = fields.Char('To', required=True)
    subject = fields.Char('Subject', required=True)
    body = fields.Html('Body')

    @api.model
    def _default_connection(self):
        return self.env['sm.365.connection'].search(
            [('state', 'in', ('outlook_connected', 'fully_connected'))],
            limit=1,
        )

    @api.onchange('partner_id')
    def _onchange_partner_id(self):
        if self.partner_id and self.partner_id.email:
            self.to_email = self.partner_id.email

    def action_send(self):
        self.ensure_one()
        if not self.connection_id:
            raise UserError(_('No active Microsoft 365 connection found.'))
        if not self.to_email:
            raise UserError(_('Recipient email is required.'))

        self.connection_id.send_email(
            self.to_email, self.subject, self.body or ''
        )

        # Post note on partner chatter
        if self.partner_id:
            self.partner_id.message_post(
                body=_(
                    'Email sent via Outlook:<br/>'
                    '<b>To:</b> %(to)s<br/>'
                    '<b>Subject:</b> %(subject)s',
                    to=self.to_email,
                    subject=self.subject,
                ),
                message_type='notification',
                subtype_xmlid='mail.mt_note',
            )

        return {
            'type': 'ir.actions.client',
            'tag': 'display_notification',
            'params': {
                'title': _('Email Sent'),
                'message': _('Email sent to %s via Outlook') % self.to_email,
                'type': 'success',
                'sticky': False,
            },
        }
