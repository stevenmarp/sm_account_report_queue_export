/** @odoo-module **/

import { FormController } from "@web/views/form/form_controller";
import { patch } from "@web/core/utils/patch";
import { useService } from "@web/core/utils/hooks";
import { onWillUnmount } from "@odoo/owl";

patch(FormController.prototype, {
    setup() {
        super.setup(...arguments);
        this.actionService = useService("action");
        if (this.props.resModel !== "sm.account.report.queue.export") {
            return;
        }
        this.smAccountReportQueueRefreshTimer = null;
        this.smAccountReportQueueRefreshing = false;
        this.smAccountReportQueueDownloaded = false;

        const refresh = async () => {
            if (this.smAccountReportQueueRefreshing) {
                return;
            }
            const root = this.model.root;
            const data = root && root.data;
            if (!root || !root.resId || !data || !["queued", "running"].includes(data.state)) {
                return;
            }
            this.smAccountReportQueueRefreshing = true;
            try {
                const wasProcessing = ["queued", "running"].includes(data.state);
                await this.model.load();
                this.render();
                const updatedData = root.data;
                if (wasProcessing && updatedData.state === "done" && !this.smAccountReportQueueDownloaded) {
                    this.smAccountReportQueueDownloaded = true;
                    await this.actionService.doAction({
                        type: "ir.actions.act_url",
                        url: `/sm_account_report_queue_export/result/${root.resId}`,
                        target: "download",
                    });
                }
            } finally {
                this.smAccountReportQueueRefreshing = false;
            }
        };

        this.smAccountReportQueueRefreshTimer = setInterval(refresh, 5000);
        onWillUnmount(() => {
            if (this.smAccountReportQueueRefreshTimer) {
                clearInterval(this.smAccountReportQueueRefreshTimer);
            }
        });
    },
});
