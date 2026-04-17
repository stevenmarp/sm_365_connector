# -*- coding: utf-8 -*-

from odoo import api, fields, models, _


class Microsoft365EmailRule(models.Model):
    _name = 'sm.365.email.rule'
    _description = 'Outlook Email Auto-Conversion Rule'
    _order = 'sequence, id'

    name = fields.Char(required=True)
    sequence = fields.Integer(default=10)
    active = fields.Boolean(default=True)
    connection_id = fields.Many2one(
        'sm.365.connection', required=True, ondelete='cascade',
    )

    # ── Matching conditions ──
    from_domain = fields.Char(
        'From Domain',
        help='Match sender email domain, e.g. "@acme.com".',
    )
    from_contains = fields.Char(
        'From Contains',
        help='Substring to match in sender address.',
    )
    subject_contains = fields.Char(
        'Subject Contains',
        help='Substring to match in subject.',
    )
    body_contains = fields.Char(
        'Body Contains',
        help='Substring to match in body preview.',
    )

    # ── Action ──
    action_type = fields.Selection([
        ('lead', 'Create CRM Lead'),
        ('opportunity', 'Create CRM Opportunity'),
        ('task', 'Create Project Task'),
        ('partner_note', 'Log on Partner'),
    ], required=True, default='lead')

    team_id = fields.Many2one('crm.team', string='Sales Team')
    user_id = fields.Many2one('res.users', string='Assign To')
    project_id = fields.Many2one('project.project', string='Project')
    tag_ids = fields.Many2many('crm.tag', string='Tags')

    match_count = fields.Integer('Matches So Far', readonly=True, default=0)

    def _match(self, msg):
        """Return True if the Outlook message dict matches this rule."""
        self.ensure_one()
        from_info = (msg.get('from') or {}).get('emailAddress') or {}
        from_email = (from_info.get('address') or '').lower()
        subject = (msg.get('subject') or '').lower()
        body = (msg.get('bodyPreview') or '').lower()

        if self.from_domain and self.from_domain.lower().lstrip('@') not in from_email:
            return False
        if self.from_contains and self.from_contains.lower() not in from_email:
            return False
        if self.subject_contains and self.subject_contains.lower() not in subject:
            return False
        if self.body_contains and self.body_contains.lower() not in body:
            return False
        return True

    def _apply(self, msg, mail_message):
        """Execute the rule's action for a matched message."""
        self.ensure_one()
        from_info = (msg.get('from') or {}).get('emailAddress') or {}
        from_email = from_info.get('address') or ''
        from_name = from_info.get('name') or from_email
        subject = msg.get('subject') or 'No Subject'
        body = msg.get('bodyPreview') or ''

        Partner = self.env['res.partner']
        partner = Partner.search([('email', '=ilike', from_email)], limit=1)
        if not partner and from_email:
            partner = Partner.create({'name': from_name, 'email': from_email})

        if self.action_type in ('lead', 'opportunity'):
            Lead = self.env['crm.lead']
            vals = {
                'name': subject,
                'description': body,
                'email_from': from_email,
                'contact_name': from_name,
                'partner_id': partner.id if partner else False,
                'type': 'lead' if self.action_type == 'lead' else 'opportunity',
                'team_id': self.team_id.id if self.team_id else False,
                'user_id': self.user_id.id if self.user_id else False,
                'tag_ids': [(6, 0, self.tag_ids.ids)] if self.tag_ids else False,
            }
            record = Lead.create({k: v for k, v in vals.items() if v is not False})
            if mail_message:
                mail_message.sudo().write({
                    'model': 'crm.lead', 'res_id': record.id,
                })
        elif self.action_type == 'task':
            Task = self.env['project.task']
            record = Task.create({
                'name': subject,
                'description': body,
                'project_id': self.project_id.id if self.project_id else False,
                'user_ids': [(6, 0, [self.user_id.id])] if self.user_id else False,
                'partner_id': partner.id if partner else False,
            })
            if mail_message:
                mail_message.sudo().write({
                    'model': 'project.task', 'res_id': record.id,
                })
        elif self.action_type == 'partner_note' and partner:
            partner.message_post(
                body=f'<b>Outlook email:</b> {subject}<br/>{body}',
                subject=subject,
            )

        self.sudo().match_count += 1
        return True
