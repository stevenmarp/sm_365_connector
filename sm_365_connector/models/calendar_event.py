# -*- coding: utf-8 -*-

from odoo import api, fields, models, _
from odoo.exceptions import UserError


class CalendarEvent(models.Model):
    _inherit = 'calendar.event'

    microsoft365_id = fields.Char(
        'Microsoft 365 ID', index=True, copy=False, readonly=True,
    )
    is_teams_meeting = fields.Boolean(
        'Teams Meeting',
        help='Create an online Microsoft Teams meeting for this event.',
    )
    teams_join_url = fields.Char('Teams Join URL', readonly=True, copy=False)
    teams_conference_id = fields.Char(
        'Teams Conference ID', readonly=True, copy=False,
    )

    def action_create_teams_meeting(self):
        """Create a Microsoft Teams online meeting for this event."""
        self.ensure_one()
        conn = self.env['sm.365.connection']._get_user_connection()
        if not conn:
            raise UserError(
                _('No active Microsoft 365 connection for the current user.')
            )
        conn._create_teams_meeting_for_event(self)
        return {
            'type': 'ir.actions.client',
            'tag': 'display_notification',
            'params': {
                'title': _('Teams Meeting'),
                'message': _('Teams meeting link created.'),
                'type': 'success',
                'sticky': False,
            },
        }

    def action_join_teams_meeting(self):
        self.ensure_one()
        if not self.teams_join_url:
            raise UserError(_('No Teams join URL. Create the meeting first.'))
        return {
            'type': 'ir.actions.act_url',
            'url': self.teams_join_url,
            'target': 'new',
        }
