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
    function (
        Controller,
        JSONModel,
        Filter,
        FilterOperator,
        MessageBox,
        MessageToast,
        Formatter,
        SimpleChart,
        ExcelExport
    ) {
        "use strict";

        const PAGE_SIZE = 5000;    // mỗi lượt request OData
        const HARD_LIMIT = 50000;  // trần an toàn phía client, vượt thì yêu cầu thu hẹp filter
        // Trạng thái tính là "đã release" (Released trở lên)
        const RELEASED_SET = { Released: true, Completed: true, Closed: true };

        return Controller.extend("zuprpt.controller.PpRpt", {
            formatter: Formatter,

            onInit() {
                this.getView().setModel(
                    new JSONModel({
                        rows: [],
                        chart: [],
                        kpi: {
                            orderCount: 0, totalQty: 0, plantCount: 0,
                            avgLeadTime: 0, releasedCount: 0,
                        },
                    }),
                    "rpt"
                );
                this.getOwnerComponent()
                    .getRouter()
                    .getRoute("RoutePpRpt")
                    .attachPatternMatched(this._onRouteMatched, this);
            },

            onExit() {
                const oBusy = this.byId("idBusyDialog");
                if (oBusy) { oBusy.close(); oBusy.destroy(); }
            },

            _onRouteMatched() {
                const oRange = this.byId("ppDateRange");
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
                this.byId("ppDateRange").setDateValue(null);
                this.byId("ppDateRange").setSecondDateValue(null);
                this.byId("ppPlant").setValue("");
                this.byId("ppOrderType").setValue("");
                this.byId("ppMaterial").setValue("");
                this.byId("ppStatus").setSelectedKey("");
            },

            _buildFilters() {
                const aFilters = [];
                const oRange = this.byId("ppDateRange");
                if (oRange.getDateValue() && oRange.getSecondDateValue()) {
                    aFilters.push(
                        new Filter("PstDate", FilterOperator.BT,
                            this._iso(oRange.getDateValue()),
                            this._iso(oRange.getSecondDateValue()))
                    );
                }
                const sPlant = this.byId("ppPlant").getValue().trim().toUpperCase();
                if (sPlant) aFilters.push(new Filter("ProductionPlant", FilterOperator.EQ, sPlant));

                const sType = this.byId("ppOrderType").getValue().trim().toUpperCase();
                if (sType) aFilters.push(new Filter("ProductionOrderType", FilterOperator.EQ, sType));

                const sMat = this.byId("ppMaterial").getValue().trim().toUpperCase();
                if (sMat) aFilters.push(new Filter("Material", FilterOperator.Contains, sMat));

                // OrderStatus là field CASE trong CDS, V4 vẫn đẩy $filter xuống được
                const sStatus = this.byId("ppStatus").getSelectedKey();
                if (sStatus) aFilters.push(new Filter("OrderStatus", FilterOperator.EQ, sStatus));

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
                        "/PPUploadReport", null, [],
                        this._buildFilters(),
                        { $orderby: "PstDate desc,ProductionOrder desc", $count: true }
                    );

                    // 1. Lấy tổng số dòng theo filter trước khi tải
                    await oBinding.requestContexts(0, 1);
                    const iTotal =
                        Number(await oBinding.getHeaderContext().requestProperty("$count")) || 0;

                    if (iTotal === 0) {
                        this._applyData([]);
                        MessageToast.show("Không có dữ liệu theo bộ lọc");
                        return;
                    }
                    if (iTotal > HARD_LIMIT) {
                        this._applyData([]);
                        MessageBox.warning(
                            `Bộ lọc hiện tại trả về ${iTotal.toLocaleString("vi-VN")} dòng, vượt giới hạn ` +
                            `${HARD_LIMIT.toLocaleString("vi-VN")} dòng của báo cáo trực tuyến. ` +
                            `Vui lòng thu hẹp khoảng thời gian hoặc thêm điều kiện lọc.`
                        );
                        return;
                    }

                    // 2. Tải tuần tự từng trang đến hết
                    const aRows = [];
                    for (let iStart = 0; iStart < iTotal; iStart += PAGE_SIZE) {
                        if (oBusy) {
                            oBusy.setText(
                                `Đang tải ${Math.min(iStart + PAGE_SIZE, iTotal).toLocaleString("vi-VN")}` +
                                ` / ${iTotal.toLocaleString("vi-VN")} dòng...`
                            );
                        }
                        const aCtx = await oBinding.requestContexts(
                            iStart, Math.min(PAGE_SIZE, iTotal - iStart)
                        );
                        aCtx.forEach((c) => aRows.push(c.getObject()));
                    }

                    this._applyData(aRows);
                    MessageToast.show(`Đã tải ${aRows.length.toLocaleString("vi-VN")} lệnh sản xuất`);
                } catch (e) {
                    MessageBox.error("Lỗi tải dữ liệu: " + (e.message || e));
                } finally {
                    if (oBusy) {
                        oBusy.setText("Đang tải dữ liệu...");
                        oBusy.close();
                    }
                }
            },

            _applyData(aRows) {
                const oRpt = this.getView().getModel("rpt");
                const oPlants = new Set();
                let fQty = 0, iLeadSum = 0, iReleased = 0;

                aRows.forEach((r) => {
                    oPlants.add(r.ProductionPlant);
                    const fRowQty = Number(r.TotalQty) || 0;
                    fQty += fRowQty;
                    iLeadSum += Number(r.LeadTimeDays) || 0;
                    if (RELEASED_SET[r.OrderStatus]) iReleased += 1;
                });

                this._aRows = aRows;
                oRpt.setProperty("/rows", aRows);
                this._renderChart();

                oRpt.setProperty("/kpi", {
                    orderCount: aRows.length,
                    totalQty: Formatter.qty(fQty),
                    plantCount: oPlants.size,
                    avgLeadTime: aRows.length
                        ? Math.round((iLeadSum / aRows.length) * 10) / 10
                        : 0,
                    releasedCount: iReleased,
                });
            },


            // ===== Chart phân tích động =====

            onChartConfigChange() {
                this._renderChart();
            },

            /** Gom nhóm client-side theo dimension/measure đang chọn, top 10 + "Khác" */
            _renderChart() {
                const oHtml = this.byId("ppChartHtml");
                if (!oHtml) return;
                const aRows = this._aRows || [];
                const sDim = this.byId("ppChartDim").getSelectedKey();
                const sMea = this.byId("ppChartMeasure").getSelectedKey();

                const fnKey = {
                    plant: (r) => r.ProductionPlant || "(trống)",
                    ordertype: (r) => r.ProductionOrderType || "(trống)",
                    material: (r) => r.Material || "(trống)",
                    status: (r) => r.OrderStatus || "N/A",
                    month: (r) => (r.PstDate ? String(r.PstDate).substring(0, 7) : "(trống)"),
                }[sDim];
                const fnVal = {
                    count: () => 1,
                    qty: (r) => Number(r.TotalQty) || 0,
                }[sMea];

                const m = {};
                aRows.forEach((r) => {
                    const k = fnKey(r);
                    m[k] = (m[k] || 0) + fnVal(r);
                });

                let aData = Object.keys(m).map((k) => ({
                    label: sDim === "month" && k.length === 7
                        ? k.substring(5, 7) + "/" + k.substring(0, 4) : k,
                    value: m[k],
                    _k: k,
                }));
                if (sDim === "month") {
                    aData.sort((a, b) => (a._k > b._k ? 1 : -1));
                } else {
                    aData.sort((a, b) => b.value - a.value);
                    if (aData.length > 10) {
                        const aTop = aData.slice(0, 10);
                        aTop.push({
                            label: "Khác",
                            value: aData.slice(10).reduce((s, d) => s + d.value, 0),
                        });
                        aData = aTop;
                    }
                }
                oHtml.setContent(SimpleChart.columns(aData));
            },

            onExport() {
                const oRpt = this.getView().getModel("rpt");
                const aRows = oRpt.getProperty("/rows");
                if (!aRows.length) {
                    MessageToast.show("Không có dữ liệu để xuất");
                    return;
                }
                const oKpi = oRpt.getProperty("/kpi");
                const oRange = this.byId("ppDateRange");

                ExcelExport.export({
                    title: "BÁO CÁO LỆNH SẢN XUẤT ĐÃ UPLOAD",
                    filePrefix: "PP_UPLOAD_REPORT",
                    filters: [
                        {
                            label: "Posting Date (log)",
                            value: oRange.getDateValue()
                                ? Formatter.date(this._iso(oRange.getDateValue())) + " - " +
                                  Formatter.date(this._iso(oRange.getSecondDateValue()))
                                : "",
                        },
                        { label: "Plant", value: this.byId("ppPlant").getValue() },
                        { label: "Order Type", value: this.byId("ppOrderType").getValue() },
                        { label: "Material chứa", value: this.byId("ppMaterial").getValue() },
                        { label: "Trạng thái lệnh", value: this.byId("ppStatus").getSelectedKey() },
                    ],
                    kpis: [
                        { label: "Tổng lệnh SX", value: oKpi.orderCount },
                        { label: "Tổng số lượng", value: oKpi.totalQty },
                        { label: "Số plant", value: oKpi.plantCount },
                        { label: "Lead time TB (ngày)", value: oKpi.avgLeadTime },
                        { label: "Đã release", value: oKpi.releasedCount },
                    ],
                    columns: [
                        { key: "ProductionOrder", label: "Production Order", width: 15 },
                        { key: "OrderStatus", label: "Trạng thái", width: 11 },
                        // { key: "ReleaseDate", label: "Release Date", width: 12, type: "date" },
                        { key: "IdDoc", label: "ID", width: 8 },
                        { key: "ProductionPlant", label: "Plant", width: 8 },
                        { key: "Material", label: "Material", width: 14 },
                        { key: "ProductionOrderType", label: "Order Type", width: 10 },
                        { key: "ProductionVersion", label: "Version", width: 9 },
                        { key: "TotalQty", label: "Qty", width: 10, type: "qty" },
                        { key: "BaseUnit", label: "Unit", width: 7 },
                        { key: "StartDate", label: "Start Date", width: 12, type: "date" },
                        { key: "EndDate", label: "End Date", width: 12, type: "date" },
                        { key: "LeadTimeDays", label: "Lead (ngày)", width: 10, type: "qty" },
                        { key: "SalesOrder", label: "Sales Order", width: 12 },
                        { key: "SalesOrderItem", label: "SO Item", width: 9 },
                        { key: "Filename", label: "Filename", width: 34 },
                        { key: "PstDate", label: "Pst Date", width: 12, type: "date" },
                        { key: "PstUser", label: "Posted By", width: 12 },
                    ],
                    rows: aRows,
                });
                MessageToast.show("Đang xuất Excel...");
            },
        });
    }
);
