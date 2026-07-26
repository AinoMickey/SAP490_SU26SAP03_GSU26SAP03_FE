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
        "./helper/GrHistoryExport",
        "./xlsx/xlsx.bundle",
    ],
    function (Controller, MessageToast, JSONModel, MessageBox, Filter, FilterOperator,
        GrExcelTemplate, ExcelParser, ApiService, GrHistoryExport) {
        "use strict";

        const UPLOAD_MODEL = "grModel";
        const POST_BTN_ID = "grPostButton";
        const CHECK_BTN_ID = "grCheckButton";
        const ACTION_FQN_RETRY = "com.sap.gateway.srvd.zmm_ui_pogr_o4.v0001.retryPost";
        const POLL_INTERVAL_MS = 5000;

        return Controller.extend("zfipkzup.controller.Gr", {
            dataUpload: [],
            _sFileName: "",
            _pollingId: null,

            onInit() {
                this.getView().setModel(new JSONModel({
                    items: [],
                    summary: { total: 0, valid: 0, invalid: 0 },
                    summaryText: "",
                    summaryState: "None",
                    validPct: 0,
                    mappingActive: false,
                    mapping: [],
                    sourceHeaders: [],
                    showCreatedByColumn: true,
                    showCreatedAtColumn: true,
                }), UPLOAD_MODEL);
                this.getView().setModel(new JSONModel({}), "grResult");
                this.onApplyHistoryFilter();
                this._startPolling();
            },


            onExit() {
                this._stopPolling();
                this._destroyBusy();
            },

            _getGrModel() {
                return this.getView().getModel("gr");
            },

            onNavBack() {
                this.getOwnerComponent().getRouter().navTo("RouteMain");
            },

            onDownloadTemplate() {
                GrExcelTemplate.download();
            },
            onToggleMappingPreview() {
                const oModel = this.getView().getModel(UPLOAD_MODEL);
                oModel.setProperty("/mappingActive", !oModel.getProperty("/mappingActive"));
            },
            onMappingChange() {
                // Mapping Engine đầy đủ (lưu/tái sử dụng mapping) chưa triển khai —
                // bảng /mapping hiện để trống nên nút này tạm ẩn, không cần xử lý gì thêm.
            },
            onCopyBatchId: function (oEvent) {
                const oContext = oEvent.getSource().getBindingContext("gr");
                const sBatchId = oContext?.getProperty("BatchId");
                if (!sBatchId) return;
                if (navigator.clipboard?.writeText) {
                    navigator.clipboard.writeText(sBatchId);
                    MessageToast.show("Đã copy Batch ID");
                }
            },
            onHistoryTableUpdateFinished() {
                // placeholder — giữ để binding updateFinished không báo lỗi thiếu handler
            },
            shortBatchId(sBatchId) {
                if (!sBatchId) return "";
                return sBatchId.length > 12 ? sBatchId.substring(0, 8) + "..." : sBatchId;
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
                    const oSheet = workbook.Sheets["Data"] || workbook.Sheets[workbook.SheetNames[0]];
                    const excelData = XLSX.utils.sheet_to_row_object_array(oSheet, { defval: "" });

                    if (!excelData || excelData.length === 0) {
                        MessageToast.show("File rỗng hoặc thiếu dữ liệu!");
                        this._refreshTable();
                        return;
                    }

                    excelData.shift(); // bỏ dòng label

                    const result = this._processRows(excelData);
                    this.dataUpload = result.data;
                    oModel.setProperty("/items", this.dataUpload);
                    oModel.setProperty("/summary", result.summary);
                    oModel.setProperty("/summaryText", result.summaryText);
                    oModel.setProperty("/summaryState", result.summaryState);
                    oModel.setProperty("/validPct", result.validPct);


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
                    const oRow = {
                        gr_number: _key(raw, "gr_number"),
                        document_date: this._convDate(_key(raw, "document_date")),
                        movement_type: _key(raw, "movement_type") || "101",
                        po_number: _key(raw, "po_number"),
                        po_item: _key(raw, "po_item"),
                        receive_qty: _key(raw, "receive_qty"),
                        unit: _key(raw, "unit"),
                        storage_location: _key(raw, "storage_location"),
                    };
                    const bOk = !!(oRow.document_date && oRow.po_number && oRow.po_item && oRow.receive_qty && oRow.unit && oRow.storage_location);
                    oRow.ValidationStatus = bOk ? "S" : "E";
                    oRow.ValidationMessage = bOk ? "" : "Thiếu PO Number/PO Item/Receive Qty/Unit/Storage Location";
                    aData.push(oRow);
                });

                const iTotal = aData.length;
                const iValid = aData.filter((r) => r.ValidationStatus === "S").length;
                const iInvalid = iTotal - iValid;

                return {
                    data: aData,
                    summary: { total: iTotal, valid: iValid, invalid: iInvalid },
                    summaryText: iInvalid === 0
                        ? `Tất cả ${iTotal} dòng hợp lệ`
                        : `${iInvalid}/${iTotal} dòng thiếu thông tin`,
                    summaryState: iInvalid === 0 ? "Success" : "Warning",
                    validPct: iTotal ? Math.round((iValid / iTotal) * 100) : 0,
                };
            },


                    _convDate(v) {
                if (!v) return v;
                if (/^\d+$/.test(String(v)) && String(v).length <= 6) {
                    const oExcelEpoch = Date.UTC(1899, 11, 30);
                    const oDate = new Date(oExcelEpoch + Number(v) * 86400000);
                    const p = (n) => String(n).padStart(2, "0");
                    return oDate.getUTCFullYear() + p(oDate.getUTCMonth() + 1) + p(oDate.getUTCDate());
                }
                if (v.length !== 10) return v;
                return `${v.substring(6)}${v.substring(3, 5)}${v.substring(0, 2)}`;
            },


            // ── Check / Post (tab Upload) ──

            onCheck() {
                this._callUpload(true);
            },

            onPost() {
                MessageBox.confirm(
                    "Post GR sẽ tạo Material Document THẬT trong SAP ngay lập tức, không thể hoàn tác dễ dàng. Tiếp tục?",
                    {
                        title: "Xác nhận Post",
                        onClose: (sAction) => {
                            if (sAction === MessageBox.Action.OK) this._callUpload(false);
                        }
                    }
                );
            },

            _callUpload: async function (bTestMode) {
                this._showBusy();
                try {
                    const aDocs = ApiService.buildGrDocs(this.dataUpload);
                    const oResult = await ApiService.callActionUploadGR(
                        this._getGrModel(), this.getCurrentFileName(), bTestMode, aDocs
                    );

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
                        MessageToast.show("Check hoàn tất — xem tab 'Chờ xử lý' để post khi sẵn sàng.");
                        this.onRefreshPending();
                    } else {
                        MessageToast.show("Đã gửi " + (oResult.total_count || 0) + " GR, đang xử lý.");
                        this._refreshPendingAndHistory();
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

            // ── Tab Chờ xử lý ──

            onRefreshPending() {
                this.byId("grPendingTable")?.getBinding("items")?.refresh();
            },

            onPostNow: function (oEvent) {
                const oContext = oEvent.getSource().getBindingContext("gr");
                if (!oContext) return;
                this._confirmAndPost([oContext]);
            },

            onPostSelected: function () {
                const aItems = this.byId("grPendingTable").getSelectedItems();
                const aContexts = aItems.map((it) => it.getBindingContext("gr"));
                if (aContexts.length === 0) {
                    MessageToast.show("Chưa chọn dòng nào.");
                    return;
                }
                this._confirmAndPost(aContexts);
            },

            _confirmAndPost: function (aContexts) {
                MessageBox.confirm(
                    `Xác nhận Post ${aContexts.length} GR đã chọn? Sẽ tạo Material Document THẬT trong SAP.`,
                    {
                        title: "Xác nhận Post",
                        onClose: async (sAction) => {
                            if (sAction !== MessageBox.Action.OK) return;
                            this._showBusy();
                            try {
                                for (const oContext of aContexts) {
                                    const oOperation = oContext.getModel().bindContext(
                                        ACTION_FQN_RETRY + "(...)", oContext
                                    );
                                    await oOperation.execute();
                                }
                                MessageToast.show("Đã gửi Post — đang xử lý nền, xem tab Lịch sử sau vài giây.");
                                this._refreshPendingAndHistory();
                            } catch (err) {
                                MessageBox.error(err?.message || "Lỗi khi Post.");
                            } finally {
                                this._closeBusy();
                            }
                        }
                    }
                );
            },

            // ── Tab Lịch sử ──

            onRefreshHistory() {
                this.byId("grHistoryTable")?.getBinding("items")?.refresh();
            },

            _refreshPendingAndHistory() {
                this.onRefreshPending();
                this.onRefreshHistory();
            },

              formatIsError(sStatus) {
                return sStatus === "E" || sStatus === "P";
            },

            onApplyHistoryFilter() {
                const aFilters = [];
                const sStatus = this.byId("grFilterStatus")?.getSelectedKey();
                const sSearch = this.byId("grFilterSearch")?.getValue();
                const sBatch = this.byId("grFilterBatch")?.getValue();

                if (sStatus) aFilters.push(new Filter("Status", FilterOperator.EQ, sStatus));
                if (sSearch) aFilters.push(new Filter("GrNumber", FilterOperator.Contains, sSearch));
                if (sBatch) aFilters.push(new Filter("BatchId", FilterOperator.Contains, sBatch));

                const oBinding = this.byId("grHistoryTable")?.getBinding("items");
                if (oBinding) oBinding.filter(aFilters);
            },
            onResetHistoryFilter() {
                this.byId("grFilterStatus").setSelectedKey("");
                this.byId("grFilterSearch").setValue("");
                this.byId("grFilterBatch").setValue("");
                this.onApplyHistoryFilter();
            },

            onRetry: async function (oEvent) {
                const oContext = oEvent.getSource().getBindingContext("gr");
                if (!oContext) return;
                this._showBusy();
                try {
                    const oOperation = oContext.getModel().bindContext(ACTION_FQN_RETRY + "(...)", oContext);
                    await oOperation.execute();
                    MessageToast.show("Đã gửi lại — đang xử lý nền.");
                    this._refreshPendingAndHistory();
                } catch (err) {
                    MessageBox.error(err?.message || "Lỗi khi Retry.");
                } finally {
                    this._closeBusy();
                }
            },

                        onRetrySelected: function () {
                const aItems = this.byId("grHistoryTable").getSelectedItems();
                const aContexts = aItems
                    .map((it) => it.getBindingContext("gr"))
                    .filter((ctx) => ctx && (ctx.getProperty("Status") === "E" || ctx.getProperty("Status") === "P"));

                if (aContexts.length === 0) {
                    MessageToast.show("Chưa chọn dòng Lỗi/Đang xử lý nào.");
                    return;
                }
                this._confirmAndPost(aContexts);
            },

            onExportHistory: async function () {
                const oBinding = this.byId("grHistoryTable")?.getBinding("items");
                if (!oBinding) return;
                const aContexts = await oBinding.requestContexts(0, 5000);
                const aRows = aContexts.map((c) => c.getObject());
                if (!aRows.length) {
                    MessageToast.show("Không có dữ liệu để xuất");
                    return;
                }
                GrHistoryExport.export(aRows);
            },

            // ── Xem chi tiết item (dùng chung cho tab Chờ xử lý và Lịch sử) ──
                       onShowItems: async function (oEvent) {
                const oContext = oEvent.getParameter
                    ? oEvent.getParameter("listItem")?.getBindingContext("gr")
                    : null;
                const oCtx = oContext || oEvent.getSource().getBindingContext("gr");
                if (!oCtx) return;
                const sGrNumber = oCtx.getProperty("GrNumber");
                if (!sGrNumber) return;

                if (!this._oItemDialog) {
                    this._oItemDialog = await this.loadFragment({ name: "zfipkzup.view.fragment.GrItemDetail" });
                }

                const oItemBinding = this.byId("grItemDetailTable")?.getBinding("items");
                let oStats = { grNumber: sGrNumber, totalItems: 0, validItems: 0, invalidItems: 0, summaryState: "None", hasErrors: false, errors: [] };

                if (oItemBinding) {
                    oItemBinding.filter([new Filter("GrNumber", FilterOperator.EQ, sGrNumber)]);
                    try {
                        const aContexts = await oItemBinding.requestContexts(0, 100);
                        const aItems = aContexts.map((c) => c.getObject());
                        const aErrors = aItems.filter((it) => it.Status === "E");
                        oStats = {
                            grNumber: sGrNumber,
                            totalItems: aItems.length,
                            validItems: aItems.length - aErrors.length,
                            invalidItems: aErrors.length,
                            summaryState: aErrors.length > 0 ? "Error" : "Success",
                            hasErrors: aErrors.length > 0,
                            errors: aErrors.map((it) => ({
                                title: `Item ${it.Item} — PO ${it.PoNumber}/${it.PoItem}`,
                                description: it.Message || "(không có message)",
                            })),
                        };
                    } catch (e) { /* giữ oStats mặc định nếu request lỗi */ }
                }

                this._oItemDialog.setModel(new JSONModel(oStats), "itemDialog");
                this._oItemDialog.open();
            },

            onCloseItemDialog() {
                this._oItemDialog?.close();
            },

            // ── Auto-poll ──
            _startPolling() {
                if (this._pollingId) return;
                this._pollingId = setInterval(() => {
                    this._refreshPendingAndHistory();
                }, POLL_INTERVAL_MS);
            },
            _stopPolling() {
                if (this._pollingId) { clearInterval(this._pollingId); this._pollingId = null; }
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
