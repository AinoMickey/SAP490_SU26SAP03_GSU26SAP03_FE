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
                        { $orderby: "DocDate desc" }
                    );
                    const aContexts = await oBinding.requestContexts(0, MAX_ROWS);
                    const aRows = aContexts.map((c) => c.getObject());
                    this._applyData(aRows);
                    MessageToast.show("Đã tải " + aRows.length + " dòng");
                } catch (e) {
                    MessageBox.error("Lỗi tải dữ liệu: " + (e.message || e));
                } finally {
                    if (oBusy) oBusy.close();
                }
            },

            _applyData(aRows) {
                const oRpt = this.getView().getModel("rpt");
                let iSuccess = 0, iError = 0, iPending = 0;
                let iSecSum = 0, iSecCount = 0;
                aRows.forEach((r) => {
                    if (r.Status === "S") iSuccess++;
                    else if (r.Status === "E") iError++;
                    else iPending++;

                    if (r.DocType === "GR" && (r.Status === "S" || r.Status === "E")) {
                        iSecSum += Number(r.ProcessingSeconds) || 0;
                        iSecCount++;
                    }
                });
                this._aRows = aRows;
                oRpt.setProperty("/rows", aRows);
                const iTotal = aRows.length;
                oRpt.setProperty("/kpi", {
                    total: iTotal,
                    success: iSuccess,
                    error: iError,
                    pending: iPending,
                    successRate: iTotal ? Math.round((iSuccess / iTotal) * 100) : 0,
                    errorRate: iTotal ? Math.round((iError / iTotal) * 100) : 0,
                    avgSeconds: iSecCount ? Math.round(iSecSum / iSecCount) : 0,
                });
                this._renderChart();
            },

            onChartConfigChange() { this._renderChart(); },

            _renderChart() {
                const oHtml = this.byId("kpiChartHtml");
                if (!oHtml) return;
                const aRows = this._aRows || [];
                const sDim = this.byId("kpiChartDim").getSelectedKey();
                const STATUS_LABEL = { S: "Thanh cong", E: "Loi", R: "Nhap", P: "Dang xu ly" };
                const fnKey = {
                    doctype: (r) => r.DocType || "?",
                    status: (r) => STATUS_LABEL[r.Status] || r.Status || "?",
                    month:  (r) => r.DocDate ? String(r.DocDate).substring(0, 7) : "(trống)",
                    user:   (r) => r.CreatedBy || "(trống)",
                }[sDim];

                const m = {};
                aRows.forEach((r) => { const k = fnKey(r); m[k] = (m[k] || 0) + 1; });
                let aData = Object.keys(m).map((k) => ({
                    label: sDim === "month" && k.length === 7
                        ? k.substring(5, 7) + "/" + k.substring(0, 4) : k,
                    value: m[k], _k: k,
                }));
                if (sDim === "month") {
                    aData.sort((a, b) => (a._k > b._k ? 1 : -1));
                } else {
                    aData.sort((a, b) => b.value - a.value);
                }
                oHtml.setContent(SimpleChart.columns(aData));
            },

            onExport() {
                const aRows = this.getView().getModel("rpt").getProperty("/rows");
                if (!aRows.length) { MessageToast.show("Không có dữ liệu để xuất"); return; }
                const oKpi = this.getView().getModel("rpt").getProperty("/kpi");
                const oRange = this.byId("kpiDateRange");
                ExcelExport.export({
                    title: "BAO CAO KPI UPLOAD (GR + FI + PP)",
                    filePrefix: "UPLOAD_KPI_REPORT",
                    filters: [
                        { label: "Ngày chứng từ", value: oRange.getDateValue()
                            ? this._iso(oRange.getDateValue()) + " - " + this._iso(oRange.getSecondDateValue()) : "" },
                        { label: "Loại chứng từ", value: this.byId("kpiDocType").getSelectedKey() },
                        { label: "Status", value: this.byId("kpiStatus").getSelectedKey() },
                        { label: "Created By", value: this.byId("kpiCreatedBy").getValue() },
                    ],
                    kpis: [
                        { label: "Tổng số lô", value: oKpi.total },
                        { label: "Thành công", value: oKpi.success },
                        { label: "Lỗi", value: oKpi.error },
                        { label: "Đang xử lý", value: oKpi.pending },
                        { label: "TG xử lý TB (giây, chỉ GR)", value: oKpi.avgSeconds },
                    ],
                    columns: [
                        { key: "DocType",   label: "Loại",         width: 8  },
                        { key: "DocId",     label: "Mã chứng từ",  width: 16 },
                        { key: "Status",    label: "Status",       width: 8  },
                        { key: "DocDate",   label: "Ngày chứng từ",width: 12, type: "date" },
                        { key: "BatchId",   label: "Batch/Filename", width: 30 },
                        { key: "Message",   label: "Message",     width: 40 },
                        { key: "CreatedBy", label: "Created By",  width: 12 },
                    ],
                    rows: aRows,
                });
                MessageToast.show("Đang xuất Excel...");
            },
        });
    }
);
