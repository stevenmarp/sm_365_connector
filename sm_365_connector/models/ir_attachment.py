# -*- coding: utf-8 -*-

from odoo import fields, models


class IrAttachment(models.Model):
    _inherit = 'ir.attachment'

    onedrive_item_id = fields.Char(
        'OneDrive Item ID', index=True, copy=False, readonly=True,
    )
    onedrive_web_url = fields.Char(
        'OneDrive Web URL', copy=False, readonly=True,
    )
    onedrive_download_url = fields.Char(
        'OneDrive Download URL', copy=False, readonly=True,
    )
