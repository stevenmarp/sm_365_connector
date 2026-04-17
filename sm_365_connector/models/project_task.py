# -*- coding: utf-8 -*-

from odoo import fields, models


class ProjectTask(models.Model):
    _inherit = 'project.task'

    microsoft365_id = fields.Char(
        'Microsoft To Do ID', index=True, copy=False, readonly=True,
    )
    microsoft365_list_id = fields.Char(
        'Microsoft To Do List ID', index=True, copy=False, readonly=True,
    )
