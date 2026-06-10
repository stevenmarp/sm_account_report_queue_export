# -*- coding: utf-8 -*-
import gc
import hashlib
import json
import logging
import os
import tempfile
import uuid
from datetime import date, datetime, timedelta
from decimal import Decimal

from odoo import _, api, fields, models
from odoo.exceptions import UserError
from odoo.tools import SQL, config as odoo_config
from odoo.tools.misc import format_date

try:
    import xlsxwriter
except ImportError:
    xlsxwriter = None

_logger = logging.getLogger(__name__)


class AccountReportQueueExportCancelled(Exception):
    pass


class SmAccountReportQueueExport(models.Model):
    _name = 'sm.account.report.queue.export'
    _description = 'Account Report Queue Export'
    _order = 'create_date desc, id desc'

    name = fields.Char(default='General Ledger Queue Export', required=True)
    export_format = fields.Selection([('xlsx', 'XLSX')], default='xlsx', required=True, readonly=True)
    state = fields.Selection([
        ('draft', 'Draft'),
        ('queued', 'Queued'),
        ('running', 'Preparing'),
        ('done', 'Ready'),
        ('cancelled', 'Cancelled'),
        ('error', 'Error'),
    ], default='draft', required=True)
    runner_type = fields.Selection([('cron', 'Internal Queue')], default='cron', readonly=True)
    priority = fields.Integer(default=10)
    max_retries = fields.Integer(default=3)
    retry_count = fields.Integer(readonly=True)
    queued_at = fields.Datetime(readonly=True)
    started_at = fields.Datetime(readonly=True)
    finished_at = fields.Datetime(readonly=True)
    next_run_at = fields.Datetime(readonly=True)
    cancel_requested = fields.Boolean(readonly=True)
    progress = fields.Float(readonly=True)
    row_count = fields.Integer(string='Journal Items', readonly=True)
    processed_count = fields.Integer(string='Prepared Items', readonly=True)
    export_file_path = fields.Char(readonly=True)
    export_file_size = fields.Integer(string='File Size', readonly=True)
    export_file_size_mb = fields.Float(string='File Size (MB)', compute='_compute_export_file_size_mb')
    message = fields.Text(readonly=True)
    report_options_json = fields.Text(readonly=True)
    report_options_hash = fields.Char(readonly=True, index=True)
    work_file_path = fields.Char(readonly=True)
    work_phase = fields.Selection([
        ('collect', 'Collecting Lines'),
        ('finalize', 'Building XLSX'),
    ], readonly=True)

    @api.depends('export_file_size')
    def _compute_export_file_size_mb(self):
        for job in self:
            job.export_file_size_mb = (job.export_file_size or 0) / (1024 * 1024)

    @api.model
    def action_create_general_ledger_export(self, options, row_count=0):
        job = self.action_queue_general_ledger_export(options, row_count=row_count)
        return job.action_open_job()

    @api.model
    def action_queue_general_ledger_export(self, options, row_count=0):
        options_hash = self._general_ledger_options_hash(options or {})
        existing = self.search([
            ('report_options_hash', '=', options_hash),
            ('create_uid', '=', self.env.uid),
            ('state', 'in', ['queued', 'running', 'done']),
        ], limit=1)
        if existing:
            existing._mark_stale_if_needed()
            if existing.state == 'done' and not existing._has_export_file():
                existing._unlink_export_artifact()
                existing = self.browse()
        if existing:
            return existing

        job = self.create({
            'name': _('General Ledger Queue Export'),
            'report_options_json': json.dumps(options or {}, default=str),
            'report_options_hash': options_hash,
            'row_count': row_count,
        })
        job._validate_export_config()
        job.write({
            'state': 'queued',
            'runner_type': 'cron',
            'progress': 0,
            'row_count': row_count or job._estimate_row_count(),
            'processed_count': 0,
            'retry_count': 0,
            'queued_at': fields.Datetime.now(),
            'started_at': False,
            'finished_at': False,
            'next_run_at': False,
            'cancel_requested': False,
            'export_file_size': 0,
            'message': _('General Ledger export queued.'),
        })
        return job

    def action_open_job(self):
        self.ensure_one()
        return {
            'type': 'ir.actions.act_window',
            'name': _('General Ledger Queue Export'),
            'res_model': self._name,
            'res_id': self.id,
            'views': [(self.env.ref('sm_account_report_queue_export.view_sm_account_report_queue_export_form').id, 'form')],
            'view_mode': 'form',
            'target': 'current',
        }

    def action_process_now(self):
        self.ensure_one()
        if self.state == 'draft':
            self.action_prepare_export()
        else:
            self._process_export_job()
        return self.action_open_job()

    def action_prepare_export(self):
        self.ensure_one()
        self._validate_export_config()
        self._unlink_export_artifact()
        self.write({
            'state': 'queued',
            'runner_type': 'cron',
            'progress': 0,
            'row_count': self._estimate_row_count(),
            'processed_count': 0,
            'retry_count': 0,
            'queued_at': fields.Datetime.now(),
            'started_at': False,
            'finished_at': False,
            'next_run_at': False,
            'cancel_requested': False,
            'export_file_size': 0,
            'work_file_path': False,
            'work_phase': False,
            'message': _('Export queued with Internal Queue.'),
        })
        return self.action_open_job()

    def action_cancel_export(self):
        for job in self:
            if job.state not in ('queued', 'running'):
                continue
            if job.state == 'queued':
                job._unlink_export_artifact()
                job.write({
                    'state': 'cancelled',
                    'cancel_requested': False,
                    'finished_at': fields.Datetime.now(),
                    'next_run_at': False,
                    'message': _('Export cancelled.'),
                })
            else:
                job.write({
                    'cancel_requested': True,
                    'message': _('Cancel requested. Current batch will stop soon.'),
                })
        return self.action_open_job() if len(self) == 1 else True

    def action_download_result(self):
        self.ensure_one()
        if not self._has_export_file():
            raise UserError(_('Export file is not available yet.'))
        return {
            'type': 'ir.actions.act_url',
            'url': '/sm_account_report_queue_export/result/%s' % self.id,
            'target': 'download',
        }

    def action_reset(self):
        for job in self:
            job._unlink_export_artifact()
            job.write({
                'state': 'draft',
                'progress': 0,
                'row_count': 0,
                'processed_count': 0,
                'export_file_size': 0,
                'retry_count': 0,
                'queued_at': False,
                'started_at': False,
                'finished_at': False,
                'next_run_at': False,
                'cancel_requested': False,
                'message': False,
                'work_file_path': False,
                'work_phase': False,
            })
        return self.action_open_job() if len(self) == 1 else True

    def _process_export_job(self):
        self.ensure_one()
        self._validate_export_config()
        self._ensure_general_ledger_cron_time_limit()
        try:
            self._raise_if_cancelled()
        except AccountReportQueueExportCancelled:
            self._mark_cancelled()
            return False

        if not self.started_at:
            self.write({
                'state': 'running',
                'started_at': fields.Datetime.now(),
                'finished_at': False,
                'next_run_at': fields.Datetime.now() + timedelta(minutes=5),
                'message': _('Preparing General Ledger export file...'),
                'progress': 0,
                'row_count': self.row_count or self._estimate_row_count(),
                'processed_count': 0,
            })
        else:
            self.write({
                'state': 'running',
                'finished_at': False,
                'next_run_at': fields.Datetime.now() + timedelta(minutes=5),
            })

        try:
            result = self._prepare_general_ledger_streaming_xlsx_file()
            if result == 'partial':
                return True
            self._raise_if_cancelled()
        except AccountReportQueueExportCancelled:
            self._mark_cancelled()
            return False
        except Exception as exc:
            _logger.exception('Account report queue export %s failed', self.id)
            self._mark_retry_or_error(exc)
            return False

        self.write({
            'state': 'done',
            'processed_count': self.processed_count or self.row_count,
            'progress': 100,
            'finished_at': fields.Datetime.now(),
            'next_run_at': False,
            'cancel_requested': False,
            'message': _('Export file is ready.'),
        })
        self.env['ir.cron']._notify_progress(done=self.processed_count or self.row_count or 1, remaining=0)
        return True

    @api.model
    def _cron_process_exports(self, limit=1):
        jobs = self._lock_queued_jobs(limit)
        for job in jobs:
            job._process_export_job()
            self.env.cr.commit()

    @api.model
    def _lock_queued_jobs(self, limit):
        self.env.cr.execute("""
            SELECT id
              FROM sm_account_report_queue_export
             WHERE runner_type = 'cron'
               AND (
                    (state = 'queued' AND (next_run_at IS NULL OR next_run_at <= NOW() AT TIME ZONE 'UTC'))
                 OR (state = 'running' AND next_run_at IS NOT NULL AND next_run_at <= NOW() AT TIME ZONE 'UTC')
                 OR (state = 'running' AND next_run_at IS NULL AND write_date <= (NOW() AT TIME ZONE 'UTC') - INTERVAL '10 minutes')
               )
             ORDER BY priority ASC, id ASC
             LIMIT %s
             FOR UPDATE SKIP LOCKED
        """, (limit,))
        return self.browse([row[0] for row in self.env.cr.fetchall()])

    def _ensure_general_ledger_cron_time_limit(self):
        if (odoo_config['limit_time_real_cron'] or 0) < 3600:
            odoo_config['limit_time_real_cron'] = 3600

    def _validate_export_config(self):
        self.ensure_one()
        if xlsxwriter is None:
            raise UserError(_('xlsxwriter Python library is required.'))
        self._general_ledger_options()

    def _mark_retry_or_error(self, exc):
        values = {
            'retry_count': self.retry_count + 1,
            'message': str(exc),
        }
        if self.retry_count + 1 <= self.max_retries:
            values.update({
                'state': 'queued',
                'next_run_at': fields.Datetime.now() + timedelta(minutes=5 * (self.retry_count + 1)),
            })
        else:
            values.update({
                'state': 'error',
                'finished_at': fields.Datetime.now(),
                'next_run_at': False,
            })
        self.write(values)

    def _mark_cancelled(self):
        self._unlink_export_artifact()
        self.write({
            'state': 'cancelled',
            'finished_at': fields.Datetime.now(),
            'next_run_at': False,
            'cancel_requested': False,
            'message': _('Export cancelled.'),
        })

    def _mark_stale_if_needed(self):
        self.ensure_one()
        if self.state == 'running' and self.write_date and self.write_date < fields.Datetime.now() - timedelta(hours=2):
            self.write({
                'state': 'queued',
                'next_run_at': False,
                'message': _('Previous worker became stale. Export queued again.'),
            })

    def _raise_if_cancelled(self):
        self.env.cr.execute('SELECT cancel_requested, state FROM sm_account_report_queue_export WHERE id = %s', (self.id,))
        row = self.env.cr.fetchone()
        if row and (row[0] or row[1] == 'cancelled'):
            raise AccountReportQueueExportCancelled()

    @api.model
    def _general_ledger_options_hash(self, options):
        payload = json.dumps(options or {}, sort_keys=True, default=str, separators=(',', ':'))
        return hashlib.sha256(('gl_queue_export_v1:xlsx:%s' % payload).encode('utf-8')).hexdigest()

    def _general_ledger_options(self):
        self.ensure_one()
        if not self.report_options_json:
            return {}
        try:
            return json.loads(self.report_options_json)
        except json.JSONDecodeError as exc:
            raise UserError(_('Invalid General Ledger options: %s') % exc) from exc

    def _general_ledger_report(self):
        options = self._general_ledger_options()
        report_id = options.get('report_id') or options.get('selected_variant_id')
        if isinstance(report_id, str) and report_id.isdigit():
            report_id = int(report_id)
        report = self.env['account.report'].browse(report_id).exists() if report_id else self.env['account.report']
        return report or self.env.ref('account_reports.general_ledger_report')

    def _general_ledger_report_options(self):
        options = self._general_ledger_options()
        report = self._general_ledger_report()
        if not options.get('columns') or not options.get('column_groups'):
            options = report.get_options(previous_options=options)
        return report, options

    def _general_ledger_print_report_options(self):
        base_options = self._general_ledger_options()
        report = self._general_ledger_report()
        options = report.get_options(previous_options={**base_options, 'export_mode': 'print'})
        options['export_mode'] = 'print'
        options['unfold_all'] = True
        return report, options

    def _general_ledger_count_query(self):
        report, options = self._general_ledger_report_options()
        queries = []
        for group_options in report._split_options_per_column_group(options).values():
            query = report._get_report_query(group_options, 'strict_range')
            queries.append(SQL(
                'SELECT COUNT(*) AS line_count FROM %(table_references)s WHERE %(search_condition)s',
                table_references=query.from_clause,
                search_condition=query.where_clause,
            ))
        return SQL(
            'SELECT COALESCE(SUM(line_count), 0) FROM (%(query)s) AS general_ledger_count',
            query=SQL(' UNION ALL ').join(SQL('(%s)', query) for query in queries),
        )

    def _estimate_row_count(self):
        self.ensure_one()
        self.env.cr.execute(self._general_ledger_count_query())
        return self.env.cr.fetchone()[0]

    def _prepare_general_ledger_streaming_xlsx_file(self):
        if not self.work_file_path or self.work_phase not in ('collect', 'finalize'):
            self._general_ledger_init_work_file()
        if self.work_phase == 'collect':
            return self._general_ledger_collect_work_chunk()
        return self._general_ledger_finalize_work_xlsx()

    def _general_ledger_init_work_file(self):
        work_path = '%s.jsonl' % self._new_export_filepath()
        os.makedirs(os.path.dirname(work_path), exist_ok=True)
        open(work_path, 'w').close()
        self.write({
            'work_file_path': work_path,
            'work_phase': 'collect',
            'processed_count': 0,
            'progress': 1,
            'message': _('Collecting General Ledger journal items...'),
        })
        self.env.cr.commit()

    def _general_ledger_collect_work_chunk(self):
        report, options = self._general_ledger_print_report_options()
        handler = self.env['account.general.ledger.report.handler']
        account_lines, _total_eval = self._general_ledger_account_lines(handler, report, options)
        processed = self.processed_count or 0
        chunk_size = self._general_ledger_chunk_size()
        aml_cursor, aml_read_cr = self._general_ledger_open_aml_cursor(
            handler, report, options, account_lines, offset=processed, limit=chunk_size
        )
        aml_iter = self._general_ledger_iter_aml_results(aml_cursor, aml_read_cr)
        chunk_count = 0
        try:
            with open(self.work_file_path, 'a', encoding='utf-8') as handle:
                for aml_result in aml_iter:
                    self._raise_if_cancelled()
                    handle.write(json.dumps(aml_result, default=self._general_ledger_json_default, ensure_ascii=False))
                    handle.write('\n')
                    chunk_count += 1
        finally:
            if aml_iter:
                aml_iter.close()

        processed += chunk_count
        done_collecting = chunk_count < chunk_size
        total = max(self.row_count or 0, processed or 1)
        values = {
            'row_count': total,
            'processed_count': processed,
            'progress': 99 if done_collecting else min(98, (processed / max(total, processed + chunk_size)) * 100),
            'next_run_at': fields.Datetime.now(),
            'message': _('Collected %(done)s/%(total)s General Ledger journal items.') % {
                'done': '{:,}'.format(processed),
                'total': '{:,}'.format(total),
            },
        }
        if done_collecting:
            values.update({
                'work_phase': 'finalize',
                'message': _('General Ledger journal items collected. Building XLSX file...'),
            })
        self.write(values)
        remaining = 1 if done_collecting else max(total - processed, 1)
        self.env['ir.cron']._notify_progress(done=processed, remaining=remaining)
        self.env.cr.commit()
        gc.collect()
        return 'partial'

    def _general_ledger_finalize_work_xlsx(self):
        if not self.work_file_path or not os.path.exists(self.work_file_path):
            raise UserError(_('General Ledger work file is missing. Please export again.'))
        work_path = self.work_file_path
        filepath = self._new_export_filepath()
        temp_path = '%s.part' % filepath
        collected_count = self.processed_count or self.row_count or 0
        processed = 0
        workbook = None
        self.write({'progress': 99, 'message': _('Building General Ledger XLSX file...')})
        self.env.cr.commit()
        try:
            workbook = xlsxwriter.Workbook(temp_path, {'constant_memory': True, 'tmpdir': os.path.dirname(temp_path)})
            report, options = self._general_ledger_print_report_options()
            handler = self.env['account.general.ledger.report.handler']
            formats = self._general_ledger_xlsx_formats(workbook)
            max_rows = 1048576
            sheet_index = 1
            sheet, row_number = self._general_ledger_add_sheet(workbook, sheet_index, report, options, formats)
            account_lines, total_eval = self._general_ledger_account_lines(handler, report, options)
            initial_balances = handler._get_initial_balance_values(
                report, [account.id for account, _line in account_lines], options
            ) if account_lines else {}

            def ensure_row():
                nonlocal sheet_index, sheet, row_number
                self._raise_if_cancelled()
                if row_number >= max_rows:
                    sheet_index += 1
                    sheet, row_number = self._general_ledger_add_sheet(workbook, sheet_index, report, options, formats)

            aml_iter = self._general_ledger_iter_work_amls()
            next_aml = next(aml_iter, None)
            try:
                for account, title_line in account_lines:
                    self._raise_if_cancelled()
                    ensure_row()
                    self._general_ledger_write_line(report, sheet, row_number, title_line, options, formats, account=account)
                    row_number += 1

                    if not next_aml or next_aml['account_id'] != account.id:
                        continue

                    parent_line_id = title_line['id']
                    init_balance_by_col_group = initial_balances.get(account.id, (account, {}))[1]
                    initial_line = report._get_partner_and_general_ledger_initial_balance_line(
                        options, parent_line_id, init_balance_by_col_group, account.currency_id
                    )
                    balance_progress = self._general_ledger_line_balance_progress(options, initial_line)

                    if initial_line:
                        ensure_row()
                        self._general_ledger_write_line(report, sheet, row_number, initial_line, options, formats)
                        row_number += 1

                    while next_aml and next_aml['account_id'] == account.id:
                        self._raise_if_cancelled()
                        eval_dict = {column_group_key: {} for column_group_key in options['column_groups']}
                        eval_dict[next_aml['column_group_key']] = next_aml
                        aml_line = handler._get_aml_line(report, parent_line_id, options, eval_dict, balance_progress)
                        ensure_row()
                        self._general_ledger_write_line(report, sheet, row_number, aml_line, options, formats)
                        row_number += 1
                        balance_progress = self._general_ledger_line_balance_progress(options, aml_line)
                        processed += 1
                        if processed % self._effective_batch_size() == 0:
                            self.write({
                                'progress': 99,
                                'message': _('Building General Ledger XLSX file... %(done)s rows written.') % {
                                    'done': '{:,}'.format(processed),
                                },
                            })
                            self.env['ir.cron']._notify_progress(done=collected_count or processed, remaining=1)
                            self.env.cr.commit()
                            gc.collect()
                        next_aml = next(aml_iter, None)

                    ensure_row()
                    self._general_ledger_write_line(report, sheet, row_number, report._generate_total_below_section_line(title_line), options, formats)
                    row_number += 1
            finally:
                aml_iter.close()

            total_line = handler._get_total_line(report, options, total_eval)
            ensure_row()
            self._general_ledger_write_line(report, sheet, row_number, total_line, options, formats)
            self.write({
                'processed_count': max(collected_count, processed),
                'row_count': max(self.row_count or 0, collected_count, processed),
                'progress': 99,
                'next_run_at': fields.Datetime.now() + timedelta(minutes=30),
                'message': _('Finalizing General Ledger XLSX exporter...'),
            })
            self.env.cr.commit()
            self._raise_if_cancelled()
            report._add_options_xlsx_sheet(workbook, [options])
            workbook.close()
            workbook = None
            self._raise_if_cancelled()
            os.replace(temp_path, filepath)
            file_size = os.path.getsize(filepath)
            if not file_size:
                raise UserError(_('General Ledger XLSX exporter returned an empty file.'))
            self.write({
                'export_file_path': filepath,
                'export_file_size': file_size,
                'processed_count': max(collected_count, processed),
                'row_count': max(self.row_count or 0, collected_count, processed),
                'work_file_path': False,
                'work_phase': False,
            })
            if os.path.exists(work_path):
                os.unlink(work_path)
        finally:
            if workbook:
                workbook.close()
            if os.path.exists(temp_path):
                os.unlink(temp_path)

    def _general_ledger_chunk_size(self):
        return int(self.env['ir.config_parameter'].sudo().get_param(
            'sm_account_report_queue_export.general_ledger_chunk_size',
            '100000',
        ))

    def _general_ledger_fetch_size(self):
        return min(max(self._effective_batch_size(), 5000), 10000)

    def _effective_batch_size(self):
        return 5000

    def _general_ledger_iter_work_amls(self):
        with open(self.work_file_path, 'r', encoding='utf-8') as handle:
            for line in handle:
                if line.strip():
                    yield json.loads(line)

    def _general_ledger_json_default(self, value):
        if isinstance(value, Decimal):
            return float(value)
        if isinstance(value, (date, datetime)):
            return value.isoformat()
        return str(value)

    def _general_ledger_xlsx_formats(self, workbook):
        base = {'font_name': 'Lato', 'font_size': 12, 'font_color': '#666666'}
        default = {**base, 'num_format': '#,##0.00'}
        text = base
        date_format = {**base, 'align': 'left', 'num_format': 'yyyy-mm-dd'}
        return {
            'title': workbook.add_format({'font_name': 'Lato', 'font_size': 12, 'bold': True, 'bottom': 2}),
            'workbook': workbook,
            'format_props': {'default': default, 'text': text, 'date': date_format},
            'levels': {
                0: {
                    'default': workbook.add_format({**default, 'bold': True, 'font_size': 13, 'bottom': 6}),
                    'text': workbook.add_format({**text, 'bold': True, 'font_size': 13, 'bottom': 6}),
                    'date': workbook.add_format({**date_format, 'bold': True, 'font_size': 13, 'bottom': 6}),
                    'total': workbook.add_format({**default, 'bold': True, 'font_size': 13, 'bottom': 6}),
                },
                1: {
                    'default': workbook.add_format({**default, 'bold': True, 'font_size': 13, 'bottom': 1}),
                    'text': workbook.add_format({**text, 'bold': True, 'font_size': 13, 'bottom': 1}),
                    'date': workbook.add_format({**date_format, 'bold': True, 'font_size': 13, 'bottom': 1}),
                    'total': workbook.add_format({**default, 'bold': True, 'font_size': 13, 'bottom': 1}),
                    'default_indent': workbook.add_format({**default, 'bold': True, 'font_size': 13, 'bottom': 1, 'indent': 1}),
                    'date_indent': workbook.add_format({**date_format, 'bold': True, 'font_size': 13, 'bottom': 1, 'indent': 1}),
                },
                2: {
                    'default': workbook.add_format({**default, 'bold': True}),
                    'text': workbook.add_format({**text, 'bold': True}),
                    'date': workbook.add_format({**date_format, 'bold': True}),
                    'initial': workbook.add_format(default),
                    'total': workbook.add_format({**default, 'bold': True}),
                    'default_indent': workbook.add_format({**default, 'bold': True, 'indent': 2}),
                    'date_indent': workbook.add_format({**date_format, 'bold': True, 'indent': 2}),
                    'initial_indent': workbook.add_format({**default, 'indent': 2}),
                    'total_indent': workbook.add_format({**default, 'bold': True, 'indent': 1}),
                },
                'default': {
                    'default': workbook.add_format(default),
                    'text': workbook.add_format(text),
                    'date': workbook.add_format(date_format),
                    'total': workbook.add_format(default),
                    'default_indent': workbook.add_format({**default, 'indent': 2}),
                    'date_indent': workbook.add_format({**date_format, 'indent': 2}),
                    'total_indent': workbook.add_format({**default, 'indent': 2}),
                },
            },
        }

    def _general_ledger_add_sheet(self, workbook, index, report, options, formats):
        sheet_name = 'General Ledger' if index == 1 else 'General Ledger %s' % index
        worksheet = workbook.add_worksheet(sheet_name[:31])
        worksheet.set_column(0, 0, 18)
        worksheet.set_column(1, 1, 67)
        worksheet.set_column(2, 8, 18)
        original_x_offset = 1
        y_offset = 0
        x_offset = original_x_offset + 1
        column_headers_render_data = report._get_column_headers_render_data(options)
        for header_level_index, header_level in enumerate(options.get('column_headers') or []):
            for header_to_render in header_level * column_headers_render_data['level_repetitions'][header_level_index]:
                colspan = header_to_render.get('colspan', column_headers_render_data['level_colspan'][header_level_index])
                colspan += 1 if options.get('show_horizontal_group_total') and header_level_index == 0 else 0
                self._general_ledger_write_cell(report, worksheet, formats, x_offset, y_offset, header_to_render.get('name', ''), formats['title'], colspan=colspan)
                x_offset += colspan
            if options.get('column_percent_comparison') == 'growth':
                self._general_ledger_write_cell(report, worksheet, formats, x_offset, y_offset, '%', formats['title'])
                x_offset += 1
            if options.get('show_horizontal_group_total') and header_level_index != 0:
                horizontal_group_name = next((group['name'] for group in options['available_horizontal_groups'] if group['id'] == options['selected_horizontal_group_id']), None)
                self._general_ledger_write_cell(report, worksheet, formats, x_offset, y_offset, horizontal_group_name, formats['title'])
                x_offset += 1
            y_offset += 1
            x_offset = original_x_offset + 1

        for subheader in column_headers_render_data['custom_subheaders']:
            colspan = subheader.get('colspan', 1)
            self._general_ledger_write_cell(report, worksheet, formats, x_offset, y_offset, subheader.get('name', ''), formats['title'], colspan=colspan)
            x_offset += colspan
        y_offset += 1
        x_offset = original_x_offset + 1

        self._general_ledger_write_cell(report, worksheet, formats, x_offset - 2, y_offset, _('Code'), formats['title'])
        self._general_ledger_write_cell(report, worksheet, formats, x_offset - 1, y_offset, _('Account Name'), formats['title'])
        for column in options['columns']:
            colspan = column.get('colspan', 1)
            self._general_ledger_write_cell(report, worksheet, formats, x_offset, y_offset, column.get('name', ''), formats['title'], colspan=colspan)
            x_offset += colspan
        if options.get('show_horizontal_group_total'):
            self._general_ledger_write_cell(report, worksheet, formats, x_offset, y_offset, options['columns'][0].get('name', ''), formats['title'], colspan=colspan)
        if options.get('column_percent_comparison') == 'growth':
            self._general_ledger_write_cell(report, worksheet, formats, x_offset, y_offset, '', formats['title'], colspan=colspan)
        return worksheet, y_offset + 1

    def _general_ledger_account_lines(self, handler, report, options):
        date_from = fields.Date.to_date(options['date']['date_from'])
        company_currency = self.env.company.currency_id
        totals_by_column_group = {key: {'debit': 0, 'credit': 0, 'balance': 0} for key in options['column_groups']}
        account_lines = []
        for account, column_group_results in handler._query_values(report, options):
            eval_dict = {}
            has_lines = False
            for column_group_key, results in column_group_results.items():
                account_sum = results.get('sum', {})
                account_un_earn = results.get('unaffected_earnings', {})
                account_debit = account_sum.get('debit', 0.0) + account_un_earn.get('debit', 0.0)
                account_credit = account_sum.get('credit', 0.0) + account_un_earn.get('credit', 0.0)
                account_balance = account_sum.get('balance', 0.0) + account_un_earn.get('balance', 0.0)
                eval_dict[column_group_key] = {
                    'amount_currency': account_sum.get('amount_currency', 0.0) + account_un_earn.get('amount_currency', 0.0),
                    'debit': account_debit,
                    'credit': account_credit,
                    'balance': account_balance,
                }
                max_date = fields.Date.to_date(account_sum.get('max_date')) if account_sum.get('max_date') else False
                has_lines = has_lines or (max_date and max_date >= date_from)
                totals_by_column_group[column_group_key]['debit'] += account_debit
                totals_by_column_group[column_group_key]['credit'] += account_credit
                totals_by_column_group[column_group_key]['balance'] += account_balance
            account_lines.append((account, handler._get_account_title_line(report, options, account, has_lines, eval_dict)))
        for totals in totals_by_column_group.values():
            totals['balance'] = company_currency.round(totals['balance'])
        return account_lines, totals_by_column_group

    def _general_ledger_open_aml_cursor(self, handler, report, options, account_lines, offset=0, limit=None):
        account_ids = [account.id for account, _title_line in account_lines]
        if not account_ids:
            return None, None
        order_values = SQL(', ').join(
            SQL('(%s, %s)', account_id, sequence)
            for sequence, account_id in enumerate(account_ids)
        )
        query = handler._get_query_amls(report, options, account_ids)
        ordered_query = SQL(
            """
                SELECT sm_gl_aml.*
                  FROM (%(query)s) AS sm_gl_aml
                  JOIN (VALUES %(order_values)s) AS account_order(account_id, sequence)
                    ON account_order.account_id = sm_gl_aml.account_id
                 ORDER BY account_order.sequence, 2, 17, 1
            """,
            query=query,
            order_values=order_values,
        )
        if limit:
            ordered_query = SQL('%s LIMIT %s OFFSET %s', ordered_query, limit, offset or 0)
        read_cr = self.pool.cursor()
        cursor = read_cr._cnx.cursor(name='sm_gl_queue_export_%s' % uuid.uuid4().hex)
        cursor.itersize = self._general_ledger_fetch_size()
        cursor.execute(ordered_query.code, ordered_query.params)
        return cursor, read_cr

    def _general_ledger_iter_aml_results(self, cursor, read_cr):
        if not cursor:
            return
        headers = self._general_ledger_aml_query_headers()
        try:
            while True:
                self._raise_if_cancelled()
                rows = cursor.fetchmany(self._general_ledger_fetch_size())
                if not rows:
                    break
                for row in rows:
                    self._raise_if_cancelled()
                    aml_result = dict(zip(headers, row))
                    if aml_result.get('ref') and aml_result.get('account_type') != 'asset_receivable':
                        aml_result['communication'] = '%s - %s' % (aml_result['ref'], aml_result.get('name') or '')
                    else:
                        aml_result['communication'] = aml_result.get('name')
                    yield aml_result
        finally:
            cursor.close()
            read_cr.close()

    def _general_ledger_aml_query_headers(self):
        return [
            'id', 'date', 'date_maturity', 'name', 'ref', 'company_id', 'account_id', 'payment_id',
            'partner_id', 'currency_id', 'amount_currency', 'invoice_date', 'date', 'debit', 'credit',
            'balance', 'move_name', 'company_currency_id', 'partner_name', 'move_type', 'account_code',
            'account_name', 'account_type', 'journal_code', 'journal_name', 'full_rec_name', 'column_group_key',
        ]

    def _general_ledger_write_cell(self, report, sheet, formats, x, y, value, style, colspan=1, datetime=False):
        if colspan == 1:
            if datetime:
                sheet.write_datetime(y, x, value, style)
            else:
                sheet.write(y, x, value, style)
        else:
            sheet.merge_range(y, x, y, x + colspan - 1, value, style)

    def _general_ledger_xlsx_format(self, formats, content_type='default', level='default'):
        workbook_formats = formats['levels']
        if isinstance(level, int) and level not in workbook_formats:
            workbook = formats['workbook']
            default = formats['format_props']['default']
            date_format = formats['format_props']['date']
            workbook_formats[level] = {
                **workbook_formats['default'],
                'default_indent': workbook.add_format({**default, 'indent': level}),
                'date_indent': workbook.add_format({**date_format, 'indent': level}),
                'total_indent': workbook.add_format({**default, 'bold': True, 'indent': level - 1}),
            }
        level_formats = workbook_formats[level]
        if '_indent' in content_type and not level_formats.get(content_type):
            return level_formats.get('default_indent', level_formats.get(content_type.removesuffix('_indent'), level_formats['default']))
        return level_formats.get(content_type, level_formats['default'])

    def _general_ledger_write_line(self, report, sheet, row, line, options, formats, account=None):
        level = line.get('level') or 0
        if line.get('level') == 0:
            row += 1
        elif not line.get('level'):
            level = 'default'
        line_id = report._parse_line_id(line.get('id')) if line.get('id') else []
        is_initial = line_id[-1][0] == 'initial' if line_id else False
        is_total = line_id[-1][0] == 'total' if line_id else line.get('name') == _('Total')
        cell_type, cell_value = report._get_cell_type_value(line)
        account_code_cell_format = self._general_ledger_xlsx_format(formats, 'text', level)
        if cell_type == 'date':
            cell_format = self._general_ledger_xlsx_format(formats, 'date_indent', level)
        elif is_initial:
            cell_format = self._general_ledger_xlsx_format(formats, 'initial_indent', level)
        elif is_total:
            cell_format = self._general_ledger_xlsx_format(formats, 'total_indent', level)
        else:
            cell_format = self._general_ledger_xlsx_format(formats, 'default_indent', level)

        if account:
            code, name = self.env['account.account']._split_code_name(line.get('name') or account.display_name)
            self._general_ledger_write_cell(report, sheet, formats, 0, row, code, account_code_cell_format)
            self._general_ledger_write_cell(report, sheet, formats, 1, row, name, cell_format)
        else:
            self._general_ledger_write_cell(report, sheet, formats, 1, row, cell_value, cell_format, datetime=cell_type == 'date')

        columns = list(line.get('columns', []))
        if options.get('column_percent_comparison') and 'column_percent_comparison_data' in line:
            columns += [line['column_percent_comparison_data']]
        if options.get('show_horizontal_group_total'):
            columns += [line.get('horizontal_group_total_data', {'name': 0})]
        for col, cell in enumerate(columns, start=2):
            cell_type, cell_value = self._general_ledger_xlsx_cell_value(report, cell)
            if cell_type == 'date':
                cell_format = self._general_ledger_xlsx_format(formats, 'date', level)
            elif is_initial:
                cell_format = self._general_ledger_xlsx_format(formats, 'initial', level)
            elif is_total:
                cell_format = self._general_ledger_xlsx_format(formats, 'total', level)
            else:
                cell_format = self._general_ledger_xlsx_format(formats, 'default', level)
            self._general_ledger_write_cell(report, sheet, formats, col + line.get('colspan', 1) - 1, row, self._xlsx_value(cell_value), cell_format, datetime=cell_type == 'date')

    def _general_ledger_xlsx_cell_value(self, report, cell):
        if not cell:
            return 'text', ''
        figure_type = cell.get('figure_type')
        if figure_type in ('date', 'datetime'):
            value = cell.get('name') if isinstance(cell.get('name'), str) else cell.get('no_format')
            return 'text', format_date(self.env, value) if value else ''
        if figure_type in ('monetary', 'float', 'integer', 'percentage') and 'no_format' in cell:
            return 'number', cell.get('no_format')
        cell_type, cell_value = report._get_cell_type_value(cell)
        if cell_type == 'date' and isinstance(cell.get('name'), str):
            return 'text', cell.get('name')
        if cell_type == 'date' and cell_value:
            return 'text', format_date(self.env, cell.get('no_format') or cell_value)
        if not cell_value and 'no_format' in cell:
            return 'text', cell.get('no_format') or ''
        return cell_type, cell_value

    def _general_ledger_line_balance_progress(self, options, line):
        progress = {column_group_key: 0 for column_group_key in options['column_groups']}
        if not line:
            return progress
        for column, cell in zip(options['columns'], line.get('columns', [])):
            if column.get('expression_label') == 'balance':
                progress[column['column_group_key']] = cell.get('no_format') or 0
        return progress

    def _export_filename(self):
        return 'general_ledger_queue_export.xlsx'

    def _new_export_filepath(self):
        directory = self._export_storage_dir()
        os.makedirs(directory, exist_ok=True)
        filename = '%s_%s' % (fields.Datetime.now().strftime('%Y%m%d%H%M%S'), self._export_filename())
        return os.path.join(directory, '%s_%s' % (uuid.uuid4().hex, filename))

    def _export_storage_dir(self):
        directory = self.env['ir.config_parameter'].sudo().get_param('sm_account_report_queue_export.export_dir')
        directory = directory or os.environ.get('SM_ACCOUNT_REPORT_QUEUE_EXPORT_DIR')
        directory = directory or os.path.join(tempfile.gettempdir(), 'odoo_account_report_queue_exports')
        return os.path.abspath(os.path.expanduser(directory))

    def _has_export_file(self):
        self.ensure_one()
        if not self.export_file_path:
            return False
        filepath = os.path.abspath(self.export_file_path)
        export_dir = self._export_storage_dir()
        if os.path.commonpath([filepath, export_dir]) != export_dir:
            return False
        return os.path.exists(filepath) and os.path.getsize(filepath) > 0

    def _unlink_export_artifact(self):
        if self.export_file_path and os.path.exists(self.export_file_path):
            os.unlink(self.export_file_path)
        if self.work_file_path and os.path.exists(self.work_file_path):
            os.unlink(self.work_file_path)
        self.export_file_path = False
        self.work_file_path = False
        self.work_phase = False

    def _export_mimetype(self):
        return 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

    def _xlsx_value(self, value):
        if isinstance(value, Decimal):
            return float(value)
        if isinstance(value, (dict, list, tuple, set)):
            return json.dumps(value, default=str, ensure_ascii=False)
        if isinstance(value, (bytes, bytearray, memoryview)):
            return bytes(value).decode('utf-8', errors='replace')
        return value
