# -*- coding: utf-8 -*-
{
    'name': 'Account Report Queue Export | General Ledger XLSX Background Export',
    'version': '19.0.1.0.0',
    'category': 'Accounting/Accounting',
    'summary': 'Export large Odoo General Ledger reports in background without browser timeout',
    'description': """
Account Report Queue Export
===========================
Queue large General Ledger XLSX exports in the background, track progress, and download the ready file from Odoo.

Key Features
------------
* Detect large General Ledger XLSX exports automatically
* Queue export jobs from the native Odoo accounting report button
* Background processing with progress, status, retry, cancel, and re-export
* XLSX output keeps Odoo General Ledger report layout
* No custom customer dependency; uses Odoo account_reports APIs
    """,
    'author': 'Steven Marp',
    'website': 'https://apps.odoo.com/apps/modules/browse?repo_maintainer_id=512936',
    'license': 'OPL-1',
    'depends': ['base', 'web', 'account', 'account_reports'],
    'data': [
        'security/security.xml',
        'security/ir.model.access.csv',
        'data/cron.xml',
        'views/account_report_queue_export_views.xml',
        'views/menu_views.xml',
    ],
    'assets': {
        'web.assets_backend': [
            'sm_account_report_queue_export/static/src/js/general_ledger_queue_export.js',
            'sm_account_report_queue_export/static/src/js/queue_export_auto_refresh.js',
            'sm_account_report_queue_export/static/src/scss/queue_export.scss',
        ],
    },
    'images': [
        'static/description/banner.gif',
        'static/description/icon.png',
    ],
    'installable': True,
    'application': False,
    'auto_install': False,
    'price': 58.29,
    'currency': 'USD',
}
