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

        return Controller.extend("zuprpt.controller.GrRpt", {
            formatter: Formatter,

            onInit() {
                this.getView().setModel(
                    new JSONModel({
                        rows: [],
                        kpi: { total: 0, success: 0, error: 0, pending: 0 },
                    }),
                    "rpt"
                );
                this.getOwnerComponent()
                    .getRouter()
                    .getRoute("RouteGrRpt")
                    .attachPatternMatched(this._onRouteMatched, this);
            },

            onExit() {
                const oBusy = this.byId("idBusyDialog");
                if (oBusy) { oBusy.close(); oBusy.destroy(); }
            },

            _onRouteMatched() {
                const oRange = this.byId("grDateRange");
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
                this.byId("grDateRange").setDateValue(null);
                this.byId("grDateRange").setSecondDateValue(null);
                this.byId("grStatus").setSelectedKey("");
                this.byId("grCreatedBy").setValue("");
            },

            _buildFilters() {
                const aFilters = [];
                const oRange = this.byId("grDateRange");
                if (oRange.getDateValue() && oRange.getSecondDateValue()) {
                    aFilters.push(new Filter("CreatedAt", FilterOperator.BT,
                        this._iso(oRange.getDateValue()) + "T00:00:00Z",
                        this._iso(oRange.getSecondDateValue()) + "T23:59:59Z"));
                }
                const sStatus = this.byId("grStatus").getSelectedKey();
                if (sStatus) aFilters.push(new Filter("Status", FilterOperator.EQ, sStatus));

                const sUser = this.byId("grCreatedBy").getValue().trim().toUpperCase();
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
                    // Dùng model "gr" trỏ zmm_ui_pogr_o4
                    const oModel = this.getView().getModel("gr");
                    const oBinding = oModel.bindList(
                        "/GrUpload", null, [],
                        this._buildFilters(),
                        { $orderby: "CreatedAt desc", $count: true }
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
                            `Hiển thị ${aRows.length}/${iTotal} GR đầu tiên`
                        );
                    } else {
                        MessageToast.show("Đã tải " + aRows.length + " GR");
                    }
                } catch (e) {
                    MessageBox.error("Lỗi tải dữ liệu: " + (e.message || e));
                } finally {
                    if (oBusy) oBusy.close();
                }
            },

            _applyData(aRows, iTotal = aRows.length) {
                const oRpt = this.getView().getModel("rpt");
                let iSuccess = 0, iError = 0, iPending = 0;
                aRows.forEach((r) => {
                    if (r.Status === "S") iSuccess++;
                    else if (r.Status === "E") iError++;
                    else iPending++;
                });
                this._aRows = aRows;
                oRpt.setProperty("/rows", aRows);
                oRpt.setProperty("/kpi", {
                    total: iTotal,
                    success: iSuccess,
                    error: iError,
                    pending: iPending,
                    successRate: iTotal ? Math.round((iSuccess * 1000) / iTotal) / 10 : 0,
                    errorRate: iTotal ? Math.round((iError * 1000) / iTotal) / 10 : 0,
                });
                this._renderChart();
            },

            onChartConfigChange() { this._renderChart(); },

            _renderChart() {
                const oHtml = this.byId("grChartHtml");
                const oDonut = this.byId("grDonutHtml");
                if (!oHtml) return;
                const aRows = this._aRows || [];
                const m = {};
                let iSuccess = 0, iError = 0, iPending = 0;
                aRows.forEach((r) => {
                    const sMonth = r.CreatedAt ? String(r.CreatedAt).substring(0, 7) : "(trống)";
                    if (!m[sMonth]) { m[sMonth] = { success: 0, error: 0, pending: 0 }; }
                    if (r.Status === "S") { m[sMonth].success++; iSuccess++; }
                    else if (r.Status === "E") { m[sMonth].error++; iError++; }
                    else { m[sMonth].pending++; iPending++; }
                });
                const aData = Object.keys(m).sort().map((k) => ({
                    label: k.length === 7 ? k.substring(5, 7) + "/" + k.substring(0, 4) : k,
                    values: [m[k].success, m[k].pending, m[k].error],
                    _k: k,
                }));
                oHtml.setContent(SimpleChart.stackedColumns(aData));
                if (oDonut) {
                    oDonut.setContent(SimpleChart.donut([
                        { label: "Thành công", value: iSuccess },
                        { label: "Đang xử lý", value: iPending },
                        { label: "Lỗi", value: iError },
                    ]));
                }
            },

            onExport() {
                const aRows = this.getView().getModel("rpt").getProperty("/rows");
                if (!aRows.length) { MessageToast.show("Không có dữ liệu để xuất"); return; }
                const oKpi = this.getView().getModel("rpt").getProperty("/kpi");
                const oRange = this.byId("grDateRange");
                ExcelExport.export({
                    title: "BÁO CÁO GR UPLOAD",
                    filePrefix: "GR_UPLOAD_REPORT",
                    filters: [
                        { label: "Created At", value: oRange.getDateValue()
                            ? this._iso(oRange.getDateValue()) + " - " + this._iso(oRange.getSecondDateValue()) : "" },
                        { label: "Status", value: this.byId("grStatus").getSelectedKey() },
                        { label: "Created By", value: this.byId("grCreatedBy").getValue() },
                    ],
                    kpis: [
                        { label: "Tổng GR", value: oKpi.total },
                        { label: "Thành công", value: oKpi.success },
                        { label: "Lỗi", value: oKpi.error },
                        { label: "Đang xử lý", value: oKpi.pending },
                    ],
                    columns: [
                        { key: "GrNumber",         label: "GR Number",      width: 14 },
                        { key: "Status",           label: "Status",         width: 8  },
                        { key: "DocumentDate",     label: "Doc. Date",      width: 12, type: "date" },
                        { key: "MaterialDocument", label: "Mat. Document",  width: 16 },
                        { key: "MatDocYear",       label: "Year",           width: 6  },
                        { key: "Message",          label: "Message",        width: 40 },
                        { key: "CreatedBy",        label: "Created By",     width: 12 },
                        { key: "CreatedAt",        label: "Created At",     width: 18, type: "date" },
                    ],
                    rows: aRows,
                });
                MessageToast.show("Đang xuất Excel...");
            },
        });
    }
);
