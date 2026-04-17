# -*- coding: utf-8 -*-

import base64
import json
import logging

from odoo import http
from odoo.http import request

_logger = logging.getLogger(__name__)


class Microsoft365AuthController(http.Controller):

    @http.route('/microsoft365/callback', type='http', auth='user', website=False)
    def oauth2_callback(self, code=None, state=None, error=None, **kw):
        """Handle the OAuth2 redirect from Microsoft."""
        if error:
            _logger.error('Microsoft OAuth2 error: %s — %s', error, kw.get('error_description'))
            return request.redirect('/odoo/action-sm_365_connector.action_365_connection')

        if not code or not state:
            _logger.error('Microsoft OAuth2 callback missing code or state')
            return request.redirect('/odoo/action-sm_365_connector.action_365_connection')

        try:
            state_data = json.loads(base64.urlsafe_b64decode(state).decode())
        except Exception:
            _logger.error('Invalid OAuth2 state parameter')
            return request.redirect('/odoo/action-sm_365_connector.action_365_connection')

        connection_id = state_data.get('cid')
        if not connection_id:
            _logger.error('No connection ID in OAuth2 state')
            return request.redirect('/odoo/action-sm_365_connector.action_365_connection')

        connection = (
            request.env['sm.365.connection']
            .sudo()
            .browse(connection_id)
            .exists()
        )
        if not connection:
            _logger.error('Connection %s not found', connection_id)
            return request.redirect('/odoo/action-sm_365_connector.action_365_connection')

        try:
            connection._exchange_code_for_token(code)
        except Exception as exc:
            _logger.error('Token exchange failed: %s', exc)

        return request.redirect(
            f'/odoo/microsoft-365-connector/{connection.id}'
        )
