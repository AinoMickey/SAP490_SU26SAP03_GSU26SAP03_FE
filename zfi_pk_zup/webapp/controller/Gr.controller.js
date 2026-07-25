sap.ui.define(
    [
        "sap/ui/core/mvc/Controller",
        "sap/m/MessageToast",
        "sap/ui/model/json/JSONModel",
        "sap/m/MessageBox",
        "sap/ui/model/Filter",
        "sap/ui/model/FilterOperator",
        "./helper/GrExcelTemplate",
        "./helper/ExcelParser",
        "./helper/ApiService",
        "./xlsx/xlsx.bundle",
    ],
    function (Controller, MessageToast, JSONModel, MessageBox, Filter, FilterOperator,
              GrExcelTemplate, ExcelParser, ApiService) {
        "use strict";

        const UPLOAD_MODEL = "grModel";
        const POST_BTN_ID  = "grPostButton";
        const CHECK_BTN_ID = "grCheckButton";
        const ACTION_FQN_RETRY = "com.sap.gateway.srvd.zmm_ui_pogr_o4.v0001.retryPost";
        const POLL_INTERVAL_MS = 5000;

        return Controller.extend("zfipkzup.controller.Gr", {
            dataUpload: [],
            _sFileName: "",
            _pollingId: null,

            onInit() {
                this.getView().setModel(new JSONModel({ items: [] }), UPLOAD_MODEL);
                this.getView().setModel(new JSONModel({}), "grResult");
                this._startPolling();
                this.onApplyHistoryFilter();
            },

            onExit() {
                this._stopPolling();
                this._destroyBusy();
            },

            _getGrModel() {
                return this.getView().getModel("gr"); // OData V4 model gr
            },

            onNavBack() {
                this.getOwnerComponent().getRouter().navTo("RouteMain");
            },

            onDownloadTemplate() {
                GrExcelTemplate.download();
            },

            // ── Upload / Parse ──

            onFileChange: async function (oEvent) {
                const oFile = oEvent.getParameter("files")?.[0];
                const oModel = this.getView().getModel(UPLOAD_MODEL);

                if (!oFile) { this._refreshTable(); return; }

                this._sFileName = oFile.name || "";
                this.dataUpload = [];
                oModel.setProperty("/items", []);
                this._showBusy();

                try {
                    const fileContent = await ExcelParser.readFile(oFile);
                    const workbook = XLSX.read(fileContent, { type: "binary" });
                    const oSheet = workbook.Sheets["Data"] ||
                        workbook.Sheets[workbook.SheetNames[0]];
                    const excelData = XLSX.utils.sheet_to_row_object_array(oSheet, {
                        defval: ""
                    });

                    if (!excelData || excelData.length === 0) {
                        MessageToast.show("File rỗng hoặc thiếu dữ liệu!");
                        this._refreshTable();
                        return;
                    }

                    // Dòng 0 là header label template — bỏ qua
                    excelData.shift();

                    const result = this._processRows(excelData);
                    this.dataUpload = result.data;
                    oModel.setProperty("/items", this.dataUpload);

                    if (this.dataUpload.length === 0) {
                        MessageToast.show("Không có dòng dữ liệu hợp lệ!");
                        this.byId(POST_BTN_ID).setEnabled(false);
                        this.byId(CHECK_BTN_ID).setEnabled(false);
                        return;
                    }

                    MessageToast.show("Đọc thành công " + this.dataUpload.length + " dòng");
                    this.byId(POST_BTN_ID).setEnabled(true);
                    this.byId(CHECK_BTN_ID).setEnabled(true);

                } catch (err) {
                    MessageBox.error(err?.message || "Lỗi đọc file.");
                    this.byId(POST_BTN_ID).setEnabled(false);
                    this.byId(CHECK_BTN_ID).setEnabled(false);
                } finally {
                    this._closeBusy();
                    this.byId("grFileUploader")?.clear();
                }
            },

            _processRows(excelData) {
                const aData = [];
                const _key = (row, k) => String(row[k] || row[k.toUpperCase()] || "").trim();

                excelData.forEach((raw) => {
                    if (!raw) return;
                    if (!_key(raw, "po_number") && !_key(raw, "gr_number")) return;

                    aData.push({
                        gr_number:       _key(raw, "gr_number"),
                        document_date:   this._convDate(_key(raw, "document_date")),
                        movement_type:   _key(raw, "movement_type") || "101",
                        po_number:       _key(raw, "po_number"),
                        po_item:         _key(raw, "po_item"),
                        receive_qty:     _key(raw, "receive_qty"),
                        unit:            _key(raw, "unit"),
                        storage_location: _key(raw, "storage_location"),
                    });
                });

                return { data: aData };
            },

            /** DD/MM/YYYY → YYYYMMDD */
            _convDate(v) {
                if (!v || v.length !== 10) return v;
                return `${v.substring(6)}${v.substring(3, 5)}${v.substring(0, 2)}`;
            },

            // ── Check / Post ──

            onCheck() {
                this._callUpload(true);
            },

            onPost() {
                this._callUpload(false);
            },

            _callUpload: async function (bTestMode) {
                this._showBusy();
                try {
                    const aDocs = ApiService.buildGrDocs(this.dataUpload);
                    const oResult = await ApiService.callActionUploadGR(
                        this._getGrModel(),
                        this.getCurrentFileName(),
                        bTestMode,
                        aDocs
                    );

                    // QUAN TRỌNG: oResult.status ở đây là kết quả VALIDATE/DRY-RUN (đồng bộ),
                    // KHÔNG phải kết quả post GR thật (post thật chạy ở job nền, bất đồng bộ).
                    const oResultModel = this.getView().getModel("grResult");
                    if (bTestMode) {
                        oResultModel.setData({
                            ...oResult,
                            statusText: oResult.status === "S" ? "Hợp lệ" :
                                        oResult.status === "E" ? "Lỗi" : "Đang xử lý",
                            statusState: oResult.status === "S" ? "Success" :
                                         oResult.status === "E" ? "Error" : "Warning",
                        });
                    } else {
                        oResultModel.setData({
                            ...oResult,
                            statusText: oResult.status === "E" ? "Lỗi" : "Đang xử lý",
                            statusState: oResult.status === "E" ? "Error" : "Warning",
                        });
                    }
                    this.byId("grResultPanel").setVisible(true);

                    if (bTestMode) {
                        MessageToast.show("Check hoàn tất.");
                    } else {
                        MessageToast.show("Đã gửi " + (oResult.total_count || 0) + " GR, đang xử lý.");
                        this.onRefreshHistory();
                    }
                } catch (err) {
                    MessageBox.error(err?.message || JSON.stringify(err));
                } finally {
                    this._closeBusy();
                }
            },

            getCurrentFileName() {
                return this._sFileName || "";
            },

            onClearUpload() {
                this._refreshTable();
                this.byId("grResultPanel").setVisible(false);
                MessageToast.show("Đã xóa dữ liệu");
            },

            onRefreshHistory() {
                const oBinding = this.byId("grHistoryTable")?.getBinding("items");
                if (oBinding) oBinding.refresh();
            },

            // ── Auto-poll khi còn GR đang ở trạng thái R (đang xử lý nền) ──
            _startPolling() {
                if (this._pollingId) return;
                this._pollingId = setInterval(() => {
                    const oBinding = this.byId("grHistoryTable")?.getBinding("items");
                    if (!oBinding || !oBinding.getCurrentContexts) return;
                    const aContexts = oBinding.getCurrentContexts();
                    const bHasPending = aContexts.some(
                        (c) => c && c.getProperty && c.getProperty("Status") === "R"
                    );
                    if (bHasPending) oBinding.refresh();
                }, POLL_INTERVAL_MS);
            },
            _stopPolling() {
                if (this._pollingId) {
                    clearInterval(this._pollingId);
                    this._pollingId = null;
                }
            },

            // ── Filter Lịch sử ──
            onApplyHistoryFilter() {
                const aFilters = [];
                const sStatus = this.byId("grFilterStatus").getSelectedKey();
                const sSearch = this.byId("grFilterSearch").getValue();
                const bShowCheck = this.byId("grFilterShowCheck").getState();

                if (sStatus) aFilters.push(new Filter("Status", FilterOperator.EQ, sStatus));
                if (sSearch) aFilters.push(new Filter("GrNumber", FilterOperator.Contains, sSearch));
                if (!bShowCheck) aFilters.push(new Filter("Testmode", FilterOperator.EQ, false));

                const oBinding = this.byId("grHistoryTable").getBinding("items");
                if (oBinding) oBinding.filter(aFilters);
            },
            onResetHistoryFilter() {
                this.byId("grFilterStatus").setSelectedKey("");
                this.byId("grFilterSearch").setValue("");
                this.byId("grFilterShowCheck").setState(false);
                this.onApplyHistoryFilter();
            },

            // ── Retry (chỉ khi Status = E) ──
            onRetry: async function (oEvent) {
                const oContext = oEvent.getSource().getBindingContext("gr");
                if (!oContext) return;
                this._showBusy();
                try {
                    const oOperation = oContext.getModel().bindContext(
                        ACTION_FQN_RETRY + "(...)", oContext
                    );
                    await oOperation.execute();
                    MessageToast.show("Đã gửi lại — đang xử lý nền, chờ vài giây rồi xem Lịch sử");
                    this.onRefreshHistory();
                } catch (err) {
                    MessageBox.error(err?.message || "Lỗi khi Retry.");
                } finally {
                    this._closeBusy();
                }
            },

            // ── Xem chi tiết item của 1 GR ──
            onShowItems: async function (oEvent) {
                const oContext = oEvent.getSource().getBindingContext("gr");
                if (!oContext) return;
                const sGrNumber = oContext.getProperty("GrNumber");

                if (!this._oItemDialog) {
                    this._oItemDialog = await this.loadFragment({
                        name: "zfipkzup.view.fragment.GrItemDetail"
                    });
                }
                this._oItemDialog.setModel(new JSONModel({ grNumber: sGrNumber }), "itemDialog");

                const oTable = sap.ui.core.Fragment.byId(this._oItemDialog.getId(), "grItemDetailTable")
                    || this._oItemDialog.getContent()[0]; // fallback nếu id khác
                const oItemBinding = (this._oItemDialog.byId
                    ? this._oItemDialog.byId("grItemDetailTable")
                    : oTable)?.getBinding("items");
                if (oItemBinding) {
                    oItemBinding.filter([new Filter("GrNumber", FilterOperator.EQ, sGrNumber)]);
                }

                this._oItemDialog.open();
            },
            onCloseItemDialog() {
                this._oItemDialog?.close();
            },

            // ── UI Utilities ──
            _showBusy() { this.byId("idBusyDialog")?.open(); },
            _closeBusy() { this.byId("idBusyDialog")?.close(); },
            _destroyBusy() {
                const o = this.byId("idBusyDialog");
                if (o) { o.close(); o.destroy(); }
            },
            _refreshTable() {
                this.getView().getModel(UPLOAD_MODEL).setProperty("/items", []);
                this.dataUpload = [];
                this._sFileName = "";
                this.byId(POST_BTN_ID).setEnabled(false);
                this.byId(CHECK_BTN_ID).setEnabled(false);
                this.byId("grFileUploader")?.clear();
            },
        });
    }
);
