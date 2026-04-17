# -*- coding: utf-8 -*-

import json
import logging

from odoo import fields, http
from odoo.http import request

_logger = logging.getLogger(__name__)


class Microsoft365WebhookController(http.Controller):
    """Microsoft Graph change notification endpoint.

    https://learn.microsoft.com/en-us/graph/webhooks
    """

    @http.route(
        '/microsoft365/webhook', type='http', auth='public',
        methods=['POST'], csrf=False, save_session=False,
    )
    def graph_webhook(self, **kw):
        # Handle the validation handshake when subscription is created
        validation_token = request.httprequest.args.get('validationToken')
        if validation_token:
            return request.make_response(
                validation_token, headers=[('Content-Type', 'text/plain')]
            )

        try:
            body = request.httprequest.get_data(as_text=True)
            payload = json.loads(body or '{}')
        except Exception:
            _logger.warning('Microsoft365 webhook: invalid JSON payload')
            return request.make_response('', status=202)

        Sub = request.env['sm.365.subscription'].sudo()
        Log = request.env['sm.365.sync.log'].sudo()

        for notif in payload.get('value', []):
            sub_id = notif.get('subscriptionId')
            client_state = notif.get('clientState')
            resource = notif.get('resource', '')
            change_type = notif.get('changeType', '')

            sub = Sub.search([('subscription_id', '=', sub_id)], limit=1)
            if not sub:
                continue
            # Validate clientState to prevent spoofing
            if sub.client_state and client_state != sub.client_state:
                _logger.warning(
                    'Microsoft365 webhook: clientState mismatch for %s', sub_id
                )
                continue

            conn = sub.connection_id
            sub.write({
                'last_notification': fields.Datetime.now(),
                'notification_count': sub.notification_count + 1,
            })
            # Dispatch a sync based on resource type
            try:
                if 'messages' in resource and conn.sync_emails:
                    conn.action_import_emails()
                elif 'contacts' in resource and conn.sync_contacts:
                    conn.action_sync_contacts()
                elif 'events' in resource and conn.sync_calendar:
                    conn.action_sync_calendar()
            except Exception as exc:
                _logger.error('Webhook sync failed: %s', exc)
                Log.create({
                    'connection_id': conn.id,
                    'sync_type': 'auto_sync',
                    'status': 'error',
                    'details': f'Webhook {change_type} {resource}: {exc}',
                })

        return request.make_response('', status=202)
