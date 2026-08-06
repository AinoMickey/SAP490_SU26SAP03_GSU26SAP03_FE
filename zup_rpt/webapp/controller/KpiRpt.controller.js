sap.ui.define(
    [
        "sap/ui/core/mvc/Controller",
        "sap/ui/model/json/JSONModel",
        "sap/ui/model/Filter",
        "sap/ui/model/FilterOperator",
        "sap/m/MessageBox",
        "sap/m/MessageToast",
        "./helper/Formatter",
        "./helper/SimpleChart",
        "./helper/ExcelExport",
        "./xlsx/xlsx.bundle",
    ],
    function (Controller, JSONModel, Filter, FilterOperator,
              MessageBox, MessageToast, Formatter, SimpleChart, ExcelExport) {
        "use strict";

        const MAX_ROWS = 5000;

        return Controller.extend("zuprpt.controller.KpiRpt", {
            formatter: Formatter,

            onInit() {
                this.getView().setModel(
                    new JSONModel({
                        rows: [],
                        kpi: { total: 0, success: 0, error: 0, pending: 0, avgSeconds: 0 },
                    }),
                    "rpt"
                );
                this.getOwnerComponent()
                    .getRouter()
                    .getRoute("RouteKpiRpt")
                    .attachPatternMatched(this._onRouteMatched, this);
            },

            onExit() {
                const oBusy = this.byId("idBusyDialog");
                if (oBusy) { oBusy.close(); oBusy.destroy(); }
            },

            _onRouteMatched() {
                const oRange = this.byId("kpiDateRange");
                if (!oRange.getDateValue()) {
                    const oTo = new Date();
                    const oFrom = new Date();
                    oFrom.setDate(oFrom.getDate() - 30);
                    oRange.setDateValue(oFrom);
                    oRange.setSecondDateValue(oTo);
                }
                this.onGo();
            },

            onNavBack() {
                this.getOwnerComponent().getRouter().navTo("RouteMain");
            },

            onClearFilters() {
                this.byId("kpiDateRange").setDateValue(null);
                this.byId("kpiDateRange").setSecondDateValue(null);
                this.byId("kpiDocType").setSelectedKey("");
                this.byId("kpiStatus").setSelectedKey("");
                this.byId("kpiCreatedBy").setValue("");
            },

            _buildFilters() {
                const aFilters = [];
                const oRange = this.byId("kpiDateRange");
                if (oRange.getDateValue() && oRange.getSecondDateValue()) {
                    aFilters.push(new Filter("DocDate", FilterOperator.BT,
                        this._iso(oRange.getDateValue()),
                        this._iso(oRange.getSecondDateValue())));
                }
                const sDocType = this.byId("kpiDocType").getSelectedKey();
                if (sDocType) aFilters.push(new Filter("DocType", FilterOperator.EQ, sDocType));

                const sStatus = this.byId("kpiStatus").getSelectedKey();
                if (sStatus) aFilters.push(new Filter("Status", FilterOperator.EQ, sStatus));

                const sUser = this.byId("kpiCreatedBy").getValue().trim().toUpperCase();
                if (sUser) aFilters.push(new Filter("CreatedBy", FilterOperator.EQ, sUser));

                return aFilters;
            },

            _iso(oDate) {
                const p = (n) => String(n).padStart(2, "0");
                return oDate.getFullYear() + "-" + p(oDate.getMonth() + 1) + "-" + p(oDate.getDate());
            },

            onGo: async function () {
                const oBusy = this.byId("idBusyDialog");
                if (oBusy) oBusy.open();
                try {
                    const oModel = this.getView().getModel();
                    const oBinding = oModel.bindList(
                        "/UploadKPI", null, [],
                        this._buildFilters(),
                        { $orderby: "DocDate desc", $count: true }
                    );

                    await oBinding.requestContexts(0, 1);
                    const iTotal =
                        Number(
                            await oBinding.getHeaderContext().requestProperty("$count")
                        ) || 0;
                    const iLoad = Math.min(iTotal, MAX_ROWS);
                    const aContexts = iLoad
                        ? await oBinding.requestContexts(0, iLoad)
                        : [];
                    const aRows = aContexts.map((c) => c.getObject());

                    this._applyData(aRows, iTotal);
                    if (iTotal > MAX_ROWS) {
                        MessageToast.show(
                            `Hiển thị ${aRows.length}/${iTotal} dòng đầu tiên`
                        );
                    } else {
                        MessageToast.show("Đã tải " + aRows.length + " dòng");
                    }
                } catch (e) {
                    MessageBox.error("Lỗi tải dữ liệu: " + (e.message || e));
                } finally {
                    if (oBusy) oBusy.close();
                }
            },

            _applyData(aRows, iTotal = aRows.length) {
                const oRpt = this.getView().getModel("rpt");
                const oSummary = {
                    GR: { total: 0, success: 0, error: 0, pending: 0 },
                    FI: { total: 0, success: 0, error: 0, pending: 0 },
                    PP: { total: 0, success: 0, error: 0, pending: 0 },
                };
                const mBatches = {};
                let iSecSum = 0, iSecCount = 0;

                aRows.forEach((r) => {
                    const sDocType = String(r.DocType || "").toUpperCase();
                    const oType = oSummary[sDocType] || null;
                    if (oType) {
                        oType.total++;
                        if (r.Status === "S") oType.success++;
                        else if (r.Status === "E") oType.error++;
                        else oType.pending++;
                    }

                    if (r.DocType === "GR" && (r.Status === "S" || r.Status === "E")) {
                        iSecSum += Number(r.ProcessingSeconds) || 0;
                        iSecCount++;
                    }

                    const sBatchKey = (r.BatchId || "") + "|" + sDocType + "|" + (r.Filename || r.BatchId || "");
                    if (!mBatches[sBatchKey]) {
                        mBatches[sBatchKey] = {
                            BatchId: r.BatchId || "",
                            DocType: sDocType || "N/A",
                            FileName: r.Filename || r.BatchId || "",
                            total: 0,
                            success: 0,
                            error: 0,
                            processing: 0,
                            processingSeconds: 0,
                            CreatedBy: r.CreatedBy || "",
                            CreatedAt: r.CreatedAt || r.DocDate || "",
                        };
                    }
                    const oBatch = mBatches[sBatchKey];
                    oBatch.total++;
                    if (r.Status === "S") oBatch.success++;
                    else if (r.Status === "E") oBatch.error++;
                    else oBatch.processing++;
                    const nSec = Number(r.ProcessingSeconds);
                    if (!isNaN(nSec)) { oBatch.processingSeconds += nSec; }
                });

                const aBatches = Object.values(mBatches).map((oBatch) => ({
                    ...oBatch,
                    processingSeconds: oBatch.processingSeconds
                        ? Math.round((oBatch.processingSeconds / oBatch.total) * 10) / 10
                        : "",
                }));

                const fnRate = (o) => (o.total ? Math.round((o.success * 1000) / o.total) / 10 : 0);
                const fnError = (o) => (o.total ? Math.round((o.error * 1000) / o.total) / 10 : 0);
                const fnPending = (o) => (o.total ? Math.round((o.pending * 1000) / o.total) / 10 : 0);

                oRpt.setProperty("/rows", aRows);
                oRpt.setProperty("/summary", {
                    GR: { ...oSummary.GR, successRate: fnRate(oSummary.GR), errorRate: fnError(oSummary.GR), pendingRate: fnPending(oSummary.GR) },
                    FI: { ...oSummary.FI, successRate: fnRate(oSummary.FI), errorRate: fnError(oSummary.FI), pendingRate: fnPending(oSummary.FI) },
                    PP: { ...oSummary.PP, successRate: fnRate(oSummary.PP), errorRate: fnError(oSummary.PP), pendingRate: fnPending(oSummary.PP) },
                });
                oRpt.setProperty("/batches", aBatches);
                oRpt.setProperty("/kpi", {
                    total: iTotal,
                    success: oSummary.GR.success + oSummary.FI.success + oSummary.PP.success,
                    error: oSummary.GR.error + oSummary.FI.error + oSummary.PP.error,
                    pending: oSummary.GR.pending + oSummary.FI.pending + oSummary.PP.pending,
                    successRate: aRows.length ? Math.round(((oSummary.GR.success + oSummary.FI.success + oSummary.PP.success) * 1000) / aRows.length) / 10 : 0,
                    errorRate: aRows.length ? Math.round(((oSummary.GR.error + oSummary.FI.error + oSummary.PP.error) * 1000) / aRows.length) / 10 : 0,
                    avgSeconds: iSecCount ? Math.round(iSecSum / iSecCount) : 0,
                });
                this._renderChart();
            },

            onChartConfigChange() { this._renderChart(); },

            _renderChart() {
                const oHtml = this.byId("kpiChartHtml");
                if (!oHtml) return;
                const oRpt = this.getView().getModel("rpt").getProperty("/summary") || {};
                const aData = ["GR", "FI", "PP"].map((sDocType) => {
                    const oType = oRpt[sDocType] || { successRate: 0, errorRate: 0 };
                    return {
                        label: sDocType,
                        successRate: oType.successRate || 0,
                        errorRate: oType.errorRate || 0,
                    };
                });
                oHtml.setContent(SimpleChart.groupedColumns(aData));
            },

            onExport() {
                const aRows = this.getView().getModel("rpt").getProperty("/batches") || [];
                if (!aRows.length) { MessageToast.show("Không có dữ liệu để xuất"); return; }
                const oRange = this.byId("kpiDateRange");
                ExcelExport.export({
                    title: "BAO CAO KPI UPLOAD (GR + FI + PP)",
                    filePrefix: "UPLOAD_KPI_BATCH_REPORT",
                    filters: [
                        { label: "Ngày chứng từ", value: oRange.getDateValue()
                            ? this._iso(oRange.getDateValue()) + " - " + this._iso(oRange.getSecondDateValue()) : "" },
                        { label: "Loại chứng từ", value: this.byId("kpiDocType").getSelectedKey() },
                        { label: "Status", value: this.byId("kpiStatus").getSelectedKey() },
                        { label: "Created By", value: this.byId("kpiCreatedBy").getValue() },
                    ],
                    kpis: [],
                    columns: [
                        { key: "BatchId", label: "Batch ID", width: 18 },
                        { key: "DocType", label: "Loại", width: 8 },
                        { key: "FileName", label: "File name", width: 24 },
                        { key: "total", label: "Tổng dòng", width: 10 },
                        { key: "success", label: "Thành công", width: 10 },
                        { key: "error", label: "Lỗi", width: 10 },
                        { key: "processingSeconds", label: "TG xử lý (s)", width: 12 },
                        { key: "CreatedBy", label: "Created By", width: 12 },
                        { key: "CreatedAt", label: "Thời gian", width: 14 },
                    ],
                    rows: aRows,
                });
                MessageToast.show("Đang xuất Excel...");
            },
        });
    }
);
