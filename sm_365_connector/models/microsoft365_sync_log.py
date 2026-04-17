# -*- coding: utf-8 -*-

from odoo import fields, models


class Microsoft365SyncLog(models.Model):
    _name = 'sm.365.sync.log'
    _description = 'Microsoft 365 Sync Log'
    _order = 'create_date desc'

    connection_id = fields.Many2one(
        'sm.365.connection', required=True, ondelete='cascade',
    )
    sync_type = fields.Selection([
        ('outlook_auth', 'Outlook Auth'),
        ('dynamics_auth', 'Dynamics Auth'),
        ('contacts', 'Contacts'),
        ('calendar', 'Calendar'),
        ('emails', 'Emails'),
        ('tasks', 'Tasks'),
        ('send_email', 'Send Email'),
        ('crm_leads', 'CRM Leads'),
        ('crm_opportunities', 'CRM Opportunities'),
        ('crm_accounts', 'CRM Accounts'),
        ('crm_contacts', 'CRM Contacts'),
        ('auto_sync', 'Auto Sync'),
    ], required=True)
    status = fields.Selection([
        ('success', 'Success'),
        ('error', 'Error'),
    ], required=True)
    details = fields.Text()
