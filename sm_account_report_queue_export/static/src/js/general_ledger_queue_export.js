/** @odoo-module **/

import { download } from "@web/core/network/download";
import { rpc } from "@web/core/network/rpc";
import { registry } from "@web/core/registry";

async function normalAccountReportDownload(env, data) {
    env.services.ui.block();
    try {
        await download({ url: "/account_reports", data });
        if (data.no_closing_after_download !== true) {
            if (data.next_action) {
                env.services.action.doAction(data.next_action);
            } else {
                env.services.action.doAction({ type: "ir.actions.act_window_close" });
            }
        }
    } catch (error) {
        if (error.exceptionName === "AccountReportFileDownloadException") {
            const reportOptions = JSON.parse(data.options);
            const reportAction = await env.services.orm.call(
                "account.report",
                "open_account_report_file_download_error_wizard",
                [reportOptions.report_id, error.data.arguments[0], error.data.arguments[1]]
            );
            env.services.action.doAction(reportAction);
        } else {
            throw error;
        }
    } finally {
        env.services.ui.unblock();
    }
}

async function executeAccountReportDownload({ env, action }) {
    const data = action.data || {};
    if (data.file_generator !== "export_to_xlsx") {
        return normalAccountReportDownload(env, data);
    }

    const status = await rpc("/sm_account_report_queue_export/general_ledger/queue", {
        options: data.options,
        file_generator: data.file_generator,
    });
    if (status.queued && status.action) {
        return env.services.action.doAction(status.action);
    }
    return normalAccountReportDownload(env, data);
}

registry.category("action_handlers").add("ir_actions_account_report_download", executeAccountReportDownload, { force: true });
