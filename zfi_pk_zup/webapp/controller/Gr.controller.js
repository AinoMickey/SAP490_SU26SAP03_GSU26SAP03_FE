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
                this._loadAuth();
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
                const oBundle = this.getView().getModel("i18n").getResourceBundle();
                MessageToast.show(oBundle.getText("TEMPLATE_LOADING"));
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
                    const oBundle = this.getView().getModel("i18n").getResourceBundle();
                    MessageToast.show(oBundle.getText("COPY_BATCH_ID"));
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
                            const oBundle = that.getView().getModel("i18n").getResourceBundle();
                            MessageToast.show(oBundle.getText("CREATE_HISTORY_FILE", [oRes.file, oRes.rows]));
                        } catch (e) {
                            const oBundle = that.getView().getModel("i18n").getResourceBundle();
                            MessageBox.error(oBundle.getText("ERROR_BUILD_FILE", [e.message || e]));
                        } finally {
                            that._closeBusy();
                        }
                    },
                    function (oErr) {
                        that._closeBusy();
                        const oBundle = that.getView().getModel("i18n").getResourceBundle();
                        MessageBox.error(oBundle.getText("ERROR_LOAD_MODULE", ["HistoryFileExport.js", oErr.message || oErr]));
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
                        const oBundle = this.getView().getModel("i18n").getResourceBundle();
                        MessageToast.show(oBundle.getText("UPLOAD_NO_DATA"));
                        this._refreshTable();
                        return;
                    }

                    const UploadValidator = await new Promise((res) =>
                        sap.ui.require(["zfipkzup/controller/helper/UploadValidator"], res),
                    );

                    const aHeader = XLSX.utils.sheet_to_json(oSheet, { header: 1 })[0] || [];
                    const oTpl = UploadValidator.checkTemplate(aHeader, "GR");
                    if (!oTpl.ok) {
                        const oBundle = this.getView().getModel("i18n").getResourceBundle();
                        ErrorDialog.handleErrorDialog([{
                            type: "Error",
                            title: oBundle.getText("EXCEL_TEMPLATE_INVALID", ["GR"]),
                            description: "Thiếu cột: " + oTpl.missing.join(", ") +
                                ". Vui lòng tải Template mới và điền lại.",
                        }], this);
                        this._refreshTable();
                        return;
                    }

                    const oStrip = UploadValidator.stripLabelRow(excelData);
                    const aPrepared = UploadValidator.prepareRows(oStrip.rows, oStrip.startRow);

                    if (aPrepared.length === 0) {
                        const oBundle = this.getView().getModel("i18n").getResourceBundle();
                        MessageToast.show(oBundle.getText("UPLOAD_NO_DATA"));
                        this._refreshTable();
                        return;
                    }

                    const aSynErrors = UploadValidator.validateGR(aPrepared);
                    const result = this._processRows(aPrepared, aSynErrors, UploadValidator);

                    this.dataUpload = result.data;
                    this._aRowErrors = result.errors;
                    oModel.setProperty("/items", this.dataUpload);
                    oModel.setProperty("/summary", result.summary);
                    oModel.setProperty("/summaryText", result.summaryText);
                    oModel.setProperty("/summaryState", result.summaryState);
                    oModel.setProperty("/validPct", result.validPct);

                    if (aSynErrors.length > 0) {
                        this.byId(POST_BTN_ID).setEnabled(false);
                        this.byId(CHECK_BTN_ID).setEnabled(false);
                        ErrorDialog.handleErrorDialog(result.errors, this);
                        return;
                    }

                    const oBundle = this.getView().getModel("i18n").getResourceBundle();
                    MessageToast.show(oBundle.getText("UPLOAD_SUCCESS", [this.dataUpload.length]));
                    this.byId(POST_BTN_ID).setEnabled(true);
                    this.byId(CHECK_BTN_ID).setEnabled(true);
                } catch (err) {
                    const oBundle = this.getView().getModel("i18n").getResourceBundle();
                    MessageBox.error(this._extractError(err, oBundle.getText("UPLOAD_CANCELLED")));
                    this.byId(POST_BTN_ID).setEnabled(false);
                    this.byId(CHECK_BTN_ID).setEnabled(false);
                } finally {
                    this._closeBusy();
                    this.byId("grFileUploader")?.clear();
                }
            },

            /**
             * Ghép dữ liệu đã chuẩn hóa với danh sách lỗi cú pháp thành các dòng hiển thị.
             */
            _processRows(aPrepared, aSynErrors, UploadValidator) {
                const mErrByRow = {};
                aSynErrors.forEach((e) => {
                    (mErrByRow[e.excelRow] = mErrByRow[e.excelRow] || []).push(e);
                });

                const aData = aPrepared.map((o) => {
                    const r = o.r;
                    const oDate = UploadValidator.parseDate(r.document_date);
                    const aErrors = mErrByRow[o.rowNo] || [];
                    return {
                        rowNo: o.rowNo,
                        gr_number: String(r.gr_number || "").trim().toUpperCase(),
                        // ABAP nhận YYYYMMDD
                        document_date: oDate.valid
                            ? oDate.value.substring(6) + oDate.value.substring(3, 5) + oDate.value.substring(0, 2)
                            : "",
                        movement_type: String(r.movement_type || "").trim() || "101",
                        po_number: String(r.po_number || "").trim(),
                        po_item: String(r.po_item || "").trim(),
                        receive_qty: String(r.receive_qty || "").trim(),
                        unit: String(r.unit || "").trim(),
                        batch: String(r.batch || "").trim().toUpperCase(),
                        storage_location: String(r.storage_location || "").trim(),
                        errors: aErrors,
                        ValidationStatus: aErrors.length ? "E" : "S",
                        ValidationMessage: aErrors.map((e) => e.message).join(" · "),
                    };
                });

                const iTotal = aData.length;
                const iValid = aData.filter((r) => r.ValidationStatus === "S").length;
                const iInvalid = iTotal - iValid;

                return {
                    data: aData,
                    errors: aSynErrors.map((e) => ({
                        type: "Error",
                        title: `Dòng Excel ${e.excelRow} — cột ${e.field}`,
                        description: e.message,
                    })),
                    summary: { total: iTotal, valid: iValid, invalid: iInvalid, errorCount: aSynErrors.length },
                    summaryText: iInvalid === 0
                        ? `Tất cả ${iTotal} dòng hợp lệ`
                        : `${iInvalid}/${iTotal} dòng có lỗi — tổng ${aSynErrors.length} lỗi`,
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
                        title: `Dòng Excel ${oRow.rowNo} — cột ${e.field}`,
                        subtitle: `GR ${oRow.gr_number || "(trống)"}`,
                        description: e.message,
                    })),
                    this
                );
            },


            _parseDocDate(v) {
                const s = String(v || "").trim();
                if (!s) return { value: "", error: "Document Date để trống" };

                if (/^\d{5}$/.test(s)) {
                    return {
                        value: "", error:
                            `Document Date đang là số serial của Excel (${s}). ` +
                            `Định dạng ô về Text rồi nhập lại dạng DD/MM/YYYY.`
                    };
                }

                const m = s.match(/^(\d{2})[/-](\d{2})[/-](\d{4})$/);
                if (m) return { value: `${m[3]}${m[2]}${m[1]}`, error: "" };

                if (/^\d{8}$/.test(s)) return { value: s, error: "" };

                return {
                    value: "", error:
                        `Document Date "${s}" sai định dạng — phải là DD/MM/YYYY (vd 11/07/2026).`
                };
            },



            // ── Check / Post (tab Upload) ──

            onCheck() {
                this._callUpload(true);
            },

            onPost() {
                const oBundle = this.getView().getModel("i18n").getResourceBundle();
                MessageBox.confirm(oBundle.getText("CONFIRM_POST_TEXT_GR"), {
                    title: oBundle.getText("CONFIRM_POST_TITLE"),
                    onClose: (sAction) => {
                        if (sAction === MessageBox.Action.OK) this._callUpload(false);
                    }
                });
            },

            _callUpload: async function (bTestMode) {
                this._showBusy();
                try {
                    const aDocs = ApiService.buildGrDocs(this.dataUpload);
                    // Prevent duplicate submissions
                    this.byId(POST_BTN_ID)?.setEnabled(false);
                    this.byId(CHECK_BTN_ID)?.setEnabled(false);

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

                    const oBundle = this.getView().getModel("i18n").getResourceBundle();
                    if (bTestMode) {
                        MessageToast.show(oBundle.getText("CHECK_COMPLETE"));
                        this.onRefreshPending();
                    } else {
                        MessageToast.show(oBundle.getText("SENT_PROCESSING", [(oResult.total_count || 0), 'GR']));
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
                const oBundle = this.getView().getModel("i18n").getResourceBundle();
                MessageToast.show(oBundle.getText("CLEAR_UPLOAD"));
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
                    const oBundle = this.getView().getModel("i18n").getResourceBundle();
                    MessageToast.show(oBundle.getText("NO_ROW_SELECTED"));
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
                                    oOperation.setParameter("user_email", await ApiService.getUserEmail());   // ← THÊM
                                    await oOperation.execute();
                                }
                                const oBundle = this.getView().getModel("i18n").getResourceBundle();
                                MessageToast.show(oBundle.getText("SENT_PROCESSING", [aContexts.length, 'GR']));
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
                    oOperation.setParameter("user_email", await ApiService.getUserEmail());   // ← THÊM
                    await oOperation.execute();
                    const oBundle = this.getView().getModel("i18n").getResourceBundle();
                    MessageToast.show(oBundle.getText("RETRY_SENT"));
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
                    const oBundle = this.getView().getModel("i18n").getResourceBundle();
                    MessageToast.show(oBundle.getText("NO_ERROR_SELECTED"));
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
                    const oBundle = this.getView().getModel("i18n").getResourceBundle();
                    MessageToast.show(oBundle.getText("NO_DATA_EXPORT"));
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
                  _loadAuth: async function () {
                try {
                    const oAuth = await ApiService.loadMyAuth(
                        this.getOwnerComponent().getModel("gr")
                    );
                    if (oAuth.can_upload_gr === false) {
                        this.byId(POST_BTN_ID)?.setVisible(false);
                        this.byId(CHECK_BTN_ID)?.setVisible(false);
                         this.byId("fileUploader")?.setEnabled(false);   // ← THÊM: khóa nút chọn file
            this._bNoAuth = true; 
                    }
                } catch (e) {
                    // Không đọc được quyền thì để nguyên nút — backend vẫn chặn
                }
            },
            onShowAllErrors() {
                if (!this._aRowErrors?.length) return;
                ErrorDialog.handleErrorDialog(this._aRowErrors, this);
            },

        });
    }
);
