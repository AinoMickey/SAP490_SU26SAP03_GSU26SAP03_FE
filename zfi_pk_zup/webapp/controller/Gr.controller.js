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
        "./helper/ErrorDialog",
        "./xlsx/xlsx.bundle",
    ],
    function (Controller, MessageToast, JSONModel, MessageBox, Filter, FilterOperator,
        GrExcelTemplate, ExcelParser, ApiService, GrHistoryExport, ErrorDialog) {
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
                    summary: { total: 0, valid: 0, invalid: 0, errorCount: 0 },
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
                // UUID theo thời gian: phần đầu giống nhau ở mọi dòng, chỉ đuôi mới phân biệt
                return sBatchId.length > 10 ? "..." + sBatchId.slice(-10) : sBatchId;
            },


                        onDownloadHistoryFile: function (oEvent) {
                const oCtx = oEvent.getSource().getBindingContext("gr");
                if (!oCtx) return;
                const sBatchId = oCtx.getProperty("BatchId");
                const oGrModel = this._getGrModel();
                const that = this;

                this._showBusy();
                sap.ui.require(
                    ["zfipkzup/controller/helper/HistoryFileExport"],
                    async function (HistoryFileExport) {
                        try {
                            const oRes = await HistoryFileExport.exportGr(oGrModel, sBatchId, { withRefs: true });
                            MessageToast.show("Đã tạo " + oRes.file + " (" + oRes.rows + " dòng)");
                        } catch (e) {
                            MessageBox.error("Không dựng được file: " + (e.message || e));
                        } finally {
                            that._closeBusy();
                        }
                    },
                    function (oErr) {
                        that._closeBusy();
                        MessageBox.error("Không nạp được HistoryFileExport.js: " + (oErr.message || oErr));
                    }
                );
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

                    const bHasLabelRow = this._isLabelRow(excelData[0]);
                    if (bHasLabelRow) excelData.shift();

                    const result = this._processRows(excelData, bHasLabelRow ? 3 : 2);
                    this.dataUpload = result.data;
                    this._aRowErrors = result.errors;
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

                    if (result.errors.length > 0) {
                        this.byId(POST_BTN_ID).setEnabled(false);
                        this.byId(CHECK_BTN_ID).setEnabled(false);
                        ErrorDialog.handleErrorDialog(result.errors, this);
                        return;
                    }


                    MessageToast.show("Đọc thành công " + this.dataUpload.length + " dòng");
                    this.byId(POST_BTN_ID).setEnabled(true);
                    this.byId(CHECK_BTN_ID).setEnabled(true);
                } catch (err) {
                    MessageBox.error(this._extractError(err, "Lỗi đọc file."));
                    this.byId(POST_BTN_ID).setEnabled(false);
                    this.byId(CHECK_BTN_ID).setEnabled(false);
                } finally {
                    this._closeBusy();
                    this.byId("grFileUploader")?.clear();
                }
            },

                                 _processRows(excelData, iFirstExcelRow) {
                const aData = [];
                const _key = (row, k) => String(row[k] || row[k.toUpperCase()] || "").trim();

                excelData.forEach((raw, i) => {
                    if (!raw) return;
                    if (!_key(raw, "po_number") && !_key(raw, "gr_number")) return;

                    const oDate = this._parseDocDate(_key(raw, "document_date"));
                    const oRow = {
                        rowNo: iFirstExcelRow + i,
                        gr_number: _key(raw, "gr_number"),
                        document_date: oDate.value,
                        movement_type: _key(raw, "movement_type") || "101",
                        po_number: _key(raw, "po_number"),
                        po_item: _key(raw, "po_item"),
                        receive_qty: _key(raw, "receive_qty"),
                        unit: _key(raw, "unit"),
                        storage_location: _key(raw, "storage_location"),
                    };

                    const aErrors = this._validateRow(oRow, oDate.error);
                    oRow.errors = aErrors;
                    oRow.ValidationStatus = aErrors.length ? "E" : "S";
                    oRow.ValidationMessage = aErrors.map((e) => e.text).join(" · ");
                    aData.push(oRow);
                });

                const iTotal = aData.length;
                const iValid = aData.filter((r) => r.ValidationStatus === "S").length;
                const iInvalid = iTotal - iValid;
                const iErrorCount = aData.reduce((n, r) => n + r.errors.length, 0);

                return {
                    data: aData,
                    errors: aData.flatMap((r) => r.errors.map((e) => ({
                        type: "Error",
                        title: `Dòng ${r.rowNo} — cột ${e.column}`,
                        subtitle: `GR ${r.gr_number || "(trống)"}`,
                        description: e.text,
                    }))),
                    summary: { total: iTotal, valid: iValid, invalid: iInvalid, errorCount: iErrorCount },
                    summaryText: iInvalid === 0
                        ? `Tất cả ${iTotal} dòng hợp lệ`
                        : `${iInvalid}/${iTotal} dòng có lỗi — tổng ${iErrorCount} lỗi`,
                    summaryState: iInvalid === 0 ? "Success" : "Warning",
                    validPct: iTotal ? Math.round((iValid / iTotal) * 100) : 0,
                };
            },


            /**
             * Dòng nhãn có chữ nằm trong các cột vốn là số — dùng để phân biệt với dòng dữ liệu.
             */
            _isLabelRow(raw) {
                if (!raw) return false;
                const _v = (k) => String(raw[k] ?? raw[k.toUpperCase()] ?? "").trim();
                const sQty = _v("receive_qty");
                const sPo = _v("po_number");
                if (sQty && isNaN(Number(sQty))) return true;
                if (sPo && isNaN(Number(sPo))) return true;
                return /gr\s*number/i.test(_v("gr_number"));
            },

            /**
             * Mỗi ô sai sinh 1 lỗi riêng, ghi rõ tên cột và giá trị đang có.
             */
            _validateRow(oRow, sDateError) {
                const aErrors = [];
                const _req = (sField, sColumn, sLabel) => {
                    if (!String(oRow[sField] ?? "").trim()) {
                        aErrors.push({ column: sColumn, text: `${sLabel} để trống` });
                    }
                };

                _req("gr_number", "gr_number", "GR Number");
                _req("po_number", "po_number", "PO Number");
                _req("po_item", "po_item", "PO Item");
                _req("unit", "unit", "Unit");
                _req("storage_location", "storage_location", "Storage Location");

                if (sDateError) aErrors.push({ column: "document_date", text: sDateError });

                if (oRow.po_item && isNaN(Number(oRow.po_item))) {
                    aErrors.push({ column: "po_item", text: `PO Item "${oRow.po_item}" không phải số` });
                }

                const sQty = String(oRow.receive_qty ?? "").trim();
                if (!sQty) {
                    aErrors.push({ column: "receive_qty", text: "Receive Qty để trống" });
                } else if (isNaN(Number(sQty))) {
                    aErrors.push({ column: "receive_qty", text: `Receive Qty "${sQty}" không phải số` });
                } else if (Number(sQty) <= 0) {
                    aErrors.push({ column: "receive_qty", text: `Receive Qty phải > 0 (đang là ${sQty})` });
                }

                return aErrors;
            },

            onShowRowError(oEvent) {
                const oRow = oEvent.getSource().getBindingContext(UPLOAD_MODEL)?.getObject();
                if (!oRow?.errors?.length) return;
                ErrorDialog.handleErrorDialog(
                    oRow.errors.map((e) => ({
                        type: "Error",
                        title: `Cột ${e.column}`,
                        subtitle: `Dòng Excel ${oRow.rowNo} — GR ${oRow.gr_number || "(trống)"}`,
                        description: e.text,
                    })),
                    this
                );
            },

            _parseDocDate(v) {
                const s = String(v || "").trim();
                if (!s) return { value: "", error: "Document Date để trống" };

                if (/^\d{5}$/.test(s)) {
                    return { value: "", error:
                        `Document Date đang là số serial của Excel (${s}). ` +
                        `Định dạng ô về Text rồi nhập lại dạng DD/MM/YYYY.` };
                }

                const m = s.match(/^(\d{2})[/-](\d{2})[/-](\d{4})$/);
                if (m) return { value: `${m[3]}${m[2]}${m[1]}`, error: "" };

                if (/^\d{8}$/.test(s)) return { value: s, error: "" };

                return { value: "", error:
                    `Document Date "${s}" sai định dạng — phải là DD/MM/YYYY (vd 11/07/2026).` };
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
                    MessageBox.error(this._extractError(err, "Lỗi khi gọi SAP."));
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
                                MessageBox.error(this._extractError(err, "Lỗi khi Post."));
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
                     MessageBox.error(this._extractError(err, "Lỗi khi Retry."));
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
                this._aRowErrors = [];
            },

            _extractError(err, sFallback) {
                if (!err) return sFallback;

                const oErr = err.error || (err.cause && err.cause.error);
                if (oErr) {
                    const aParts = [];
                    if (oErr.message) aParts.push(oErr.message);
                    if (Array.isArray(oErr.details)) {
                        oErr.details.forEach((d) => {
                            if (d.message && d.message !== oErr.message) aParts.push(d.message);
                        });
                    }
                    if (aParts.length) return aParts.join("\n");
                }

                if (err.responseText) {
                    try {
                        const m = JSON.parse(err.responseText).error?.message;
                        if (m) return m.value || m;
                    } catch (e) {
                        return err.responseText.substring(0, 500);
                    }
                }

                return err.message || sFallback;
            },
                        onShowAllErrors() {
                if (!this._aRowErrors?.length) return;
                ErrorDialog.handleErrorDialog(this._aRowErrors, this);
            },

        });
    }
);
