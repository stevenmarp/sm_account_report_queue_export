# -*- coding: utf-8 -*-
import json

from odoo import _, models
from odoo.exceptions import RedirectWarning, UserError


class AccountReport(models.Model):
    _inherit = 'account.report'

    def export_to_xlsx(self, options, response=None):
        self.ensure_one()
        if self.env.context.get('sm_account_report_queue_export_skip_redirect'):
            return super().export_to_xlsx(options, response=response)
        row_count = self._sm_queue_general_ledger_count(options) if self._sm_queue_is_general_ledger() else 0
        if row_count and self._sm_queue_should_redirect_general_ledger(row_count):
            if not self.env.user.has_group('sm_account_report_queue_export.group_account_report_queue_export_user'):
                raise UserError(_(
                    'This General Ledger export is too large for the standard XLSX exporter. '
                    'Ask an administrator to grant Account Report Queue Export access.'
                ))
            action = self.env['sm.account.report.queue.export'].action_create_general_ledger_export(
                options,
                row_count=row_count,
            )
            raise RedirectWarning(
                _(
                    'This General Ledger export contains %(count)s rows. '
                    'Prepare it in the background to avoid browser timeout.'
                ) % {'count': row_count},
                action,
                _('Open Queue Export'),
            )
        return super().export_to_xlsx(options, response=response)

    def _sm_queue_is_general_ledger(self):
        handler = self.custom_handler_model_name or ''
        return handler == 'account.general.ledger.report.handler' or (self.name or '').lower() == 'general ledger'

    def _sm_queue_should_redirect_general_ledger(self, count):
        threshold = int(self.env['ir.config_parameter'].sudo().get_param(
            'sm_account_report_queue_export.general_ledger_xlsx_threshold',
            '100000',
        ))
        return count > threshold

    def _sm_queue_general_ledger_count(self, options):
        job = self.env['sm.account.report.queue.export'].new({
            'report_options_json': json.dumps(options or {}, default=str),
        })
        return job._estimate_row_count()
