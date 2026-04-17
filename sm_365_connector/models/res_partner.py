# -*- coding: utf-8 -*-

from odoo import fields, models


class ResPartner(models.Model):
    _inherit = 'res.partner'

    microsoft365_id = fields.Char(
        'Microsoft 365 ID', index=True, copy=False, readonly=True,
    )
    dynamics365_id = fields.Char(
        'Dynamics 365 ID', index=True, copy=False, readonly=True,
    )
