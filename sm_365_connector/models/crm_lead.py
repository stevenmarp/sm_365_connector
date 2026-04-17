# -*- coding: utf-8 -*-

from odoo import fields, models


class CrmLead(models.Model):
    _inherit = 'crm.lead'

    dynamics365_id = fields.Char(
        'Dynamics 365 ID', index=True, copy=False, readonly=True,
    )
