# -*- coding: utf-8 -*-
import json
import os

from odoo import http
from odoo.exceptions import AccessError
from odoo.http import Response, content_disposition, request


class SmAccountReportQueueExportController(http.Controller):

    def _check_access(self):
        if not (
            request.env.user.has_group('sm_account_report_queue_export.group_account_report_queue_export_user')
            or request.env.user.has_group('account.group_account_readonly')
        ):
            raise AccessError('Account Report Queue Export access denied.')

    @http.route('/sm_account_report_queue_export/general_ledger/queue', type='json', auth='user')
    def general_ledger_queue(self, options=None, file_generator=None, **kwargs):
        if file_generator != 'export_to_xlsx':
            return {'queued': False}
        self._check_access()

        options = json.loads(options or '{}') if isinstance(options, str) else (options or {})
        report_id = options.get('report_id') or options.get('selected_variant_id')
        if isinstance(report_id, str) and report_id.isdigit():
            report_id = int(report_id)
        if not report_id:
            return {'queued': False}

        allowed_company_ids = request.env['account.report'].get_report_company_ids(options)
        if not allowed_company_ids:
            company_str = request.httprequest.cookies.get('cids', str(request.env.user.company_id.id))
            allowed_company_ids = [int(str_id) for str_id in company_str.split('-')]

        report = request.env['account.report'].with_context(allowed_company_ids=allowed_company_ids).browse(report_id).exists()
        if not report or not report._sm_queue_is_general_ledger():
            return {'queued': False}

        row_count = report._sm_queue_general_ledger_count(options)
        if not row_count or not report._sm_queue_should_redirect_general_ledger(row_count):
            return {'queued': False}

        action = request.env['sm.account.report.queue.export'].action_create_general_ledger_export(
            options,
            row_count=row_count,
        )
        return {
            'queued': True,
            'action': action,
        }

    @http.route('/sm_account_report_queue_export/general_ledger/status/<int:job_id>', type='json', auth='user')
    def general_ledger_status(self, job_id, **kwargs):
        self._check_access()
        job = request.env['sm.account.report.queue.export'].browse(job_id).exists()
        if not job:
            return {'exists': False}
        job._mark_stale_if_needed()
        return self._job_payload(job, queued=True)

    @http.route('/sm_account_report_queue_export/general_ledger/cancel/<int:job_id>', type='json', auth='user')
    def general_ledger_cancel(self, job_id, **kwargs):
        self._check_access()
        job = request.env['sm.account.report.queue.export'].browse(job_id).exists()
        if not job:
            return {'exists': False}
        job.action_cancel_export()
        return self._job_payload(job, queued=True)

    def _job_payload(self, job, queued=False):
        download_url = False
        if job.state == 'done' and job._has_export_file():
            download_url = '/sm_account_report_queue_export/result/%s' % job.id
        return {
            'exists': True,
            'queued': queued,
            'id': job.id,
            'state': job.state,
            'progress': job.progress or 0,
            'row_count': job.row_count or 0,
            'processed_count': job.processed_count or 0,
            'file_size_mb': job.export_file_size_mb or 0,
            'cancel_requested': bool(job.cancel_requested),
            'message': job.message or '',
            'download_url': download_url,
        }

    @http.route('/sm_account_report_queue_export/result/<int:job_id>', type='http', auth='user', methods=['GET', 'POST'])
    def result(self, job_id, **kwargs):
        self._check_access()
        job = request.env['sm.account.report.queue.export'].browse(job_id).exists()
        if not job or not job.export_file_path:
            return request.not_found()

        filepath = os.path.abspath(job.export_file_path)
        export_dir = job._export_storage_dir()
        if os.path.commonpath([filepath, export_dir]) != export_dir or not os.path.exists(filepath) or not os.path.getsize(filepath):
            return request.not_found()

        def stream_file():
            with open(filepath, 'rb') as stream:
                while True:
                    chunk = stream.read(1024 * 1024)
                    if not chunk:
                        break
                    yield chunk

        return Response(
            stream_file(),
            headers=[
                ('Content-Type', job._export_mimetype()),
                ('Content-Disposition', content_disposition(job._export_filename())),
                ('Content-Length', str(os.path.getsize(filepath))),
            ],
            direct_passthrough=True,
        )
