# -*- coding: utf-8 -*-

from odoo import fields, models


class MailMessage(models.Model):
    _inherit = 'mail.message'

    microsoft365_id = fields.Char(
        'Microsoft 365 ID', index=True, copy=False, readonly=True,
    )
    microsoft365_conversation_id = fields.Char(
        'Outlook Conversation ID', index=True, readonly=True,
    )
    microsoft365_is_read = fields.Boolean('Outlook: Read', readonly=True)
    microsoft365_has_attachments = fields.Boolean(
        'Outlook: Has Attachments', readonly=True,
    )
