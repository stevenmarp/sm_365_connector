# -*- coding: utf-8 -*-
"""Override mail.mail to send emails via Microsoft Graph API."""

import logging

from odoo import models

_logger = logging.getLogger(__name__)


class MailMail(models.Model):
    _inherit = 'mail.mail'

    def send(self, auto_commit=False, raise_exception=False, post_send_callback=None):
        """Override send() to route emails through Microsoft Graph API
        *before* Odoo attempts the SMTP connection.
        Falls back to standard SMTP if no M365 connection is available."""
        _logger.info('=== SM365 mail.mail.send() called for %d mails ===', len(self))

        conn = self._get_m365_connection()
        if not conn:
            _logger.info('No M365 connection, falling back to SMTP')
            return super().send(
                auto_commit=auto_commit,
                raise_exception=raise_exception,
                post_send_callback=post_send_callback,
            )

        _logger.info('Sending %d mails via M365 Graph API', len(self))
        sent_ids = []
        for mail in self:
            if mail.state != 'outgoing':
                continue
            try:
                mail._send_via_graph(conn)
                mail.write({'state': 'sent', 'failure_reason': False})
                sent_ids.append(mail.id)
                if auto_commit:
                    self.env.cr.commit()
            except Exception as exc:
                _logger.warning(
                    'M365 Graph send failed for mail %s: %s', mail.id, exc,
                )
                mail.write({
                    'state': 'exception',
                    'failure_reason': str(exc)[:500],
                })
                if auto_commit:
                    self.env.cr.commit()
                if raise_exception:
                    raise

        if post_send_callback and sent_ids:
            post_send_callback(sent_ids)

        return True

    def _get_m365_connection(self):
        """Return the first active shared M365 connection with Outlook token."""
        try:
            Connection = self.env['sm.365.connection'].sudo()
            conn = Connection.search([
                ('outlook_refresh_token', '!=', False),
            ], limit=1)
            return conn or None
        except Exception:
            return None

    def _send_via_graph(self, conn):
        """Send a single mail.mail record via Microsoft Graph /me/sendMail."""
        self.ensure_one()
        conn._ensure_outlook_token()

        to_list = []
        cc_list = []

        # email_to: comma-separated email addresses
        for addr in (self.email_to or '').split(','):
            addr = addr.strip()
            if addr:
                email = self._extract_email(addr)
                to_list.append({'emailAddress': {'address': email}})

        # partner recipients
        existing = {r['emailAddress']['address'] for r in to_list}
        for partner in self.recipient_ids:
            if partner.email and partner.email not in existing:
                to_list.append({'emailAddress': {
                    'address': partner.email,
                    'name': partner.name or '',
                }})

        # CC
        if self.email_cc:
            for addr in self.email_cc.split(','):
                addr = addr.strip()
                if addr:
                    email = self._extract_email(addr)
                    cc_list.append({'emailAddress': {'address': email}})

        if not to_list:
            raise ValueError('No recipients for email %s' % self.id)

        message = {
            'subject': self.subject or '(No Subject)',
            'body': {
                'contentType': 'HTML',
                'content': self.body_html or self.body or '',
            },
            'toRecipients': to_list,
        }
        if cc_list:
            message['ccRecipients'] = cc_list

        # Attachments
        attachments = []
        for att in self.attachment_ids:
            if att.datas:
                attachments.append({
                    '@odata.type': '#microsoft.graph.fileAttachment',
                    'name': att.name or 'attachment',
                    'contentType': att.mimetype or 'application/octet-stream',
                    'contentBytes': att.datas.decode('utf-8')
                        if isinstance(att.datas, bytes)
                        else att.datas,
                })
        if attachments:
            message['attachments'] = attachments

        payload = {
            'message': message,
            'saveToSentItems': True,
        }

        conn._call_graph_api('post', '/me/sendMail', data=payload)
        _logger.info(
            'Email "%s" (id=%s) sent via M365 Graph to %s',
            self.subject, self.id,
            ', '.join(r['emailAddress']['address'] for r in to_list),
        )

    @staticmethod
    def _extract_email(addr_str):
        """Extract email from 'Name <email>' or plain 'email' string."""
        addr_str = addr_str.strip()
        if '<' in addr_str and '>' in addr_str:
            return addr_str.split('<')[1].split('>')[0].strip()
        return addr_str
