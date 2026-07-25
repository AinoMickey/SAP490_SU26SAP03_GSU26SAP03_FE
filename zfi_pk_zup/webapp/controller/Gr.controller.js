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
        const POST_BTN_ID  = "grPostButton";
        const CHECK_BTN_ID = "grCheckButton";
        const ACTION_FQN_RETRY = "com.sap.gateway.srvd.zmm_ui_pogr_o4.v0001.retryPost";
        const POLL_INTERVAL_MS = 5000;

        const GR_FIELD_MAPPING = [
            { targetField: "gr_number", label: "GR Number", required: true },
            { targetField: "document_date", label: "Document Date", required: true },
            { targetField: "movement_type", label: "Movement Type", required: false },
            { targetField: "po_number", label: "PO Number", required: true },
            { targetField: "po_item", label: "PO Item", required: true },
            { targetField: "receive_qty", label: "Receive Qty", required: true },
            { targetField: "unit", label: "Unit", required: true },
            { targetField: "storage_location", label: "Storage Location", required: true },
        ];

        return Controller.extend("zfipkzup.controller.Gr", {
            dataUpload: [],
            _sFileName: "",
            _rawExcelData: null,
            _pollingId: null,

            onInit() {
                this.getView().setModel(new JSONModel({
                    items: [],
                    mapping: [],
                    mappingActive: false,
                    summary: {
                        total: 0,
                        valid: 0,
                        invalid: 0,
                        successRate: 0,
                        errorRate: 0,
                    },
                    summaryText: "Chưa có dữ liệu",
                    summaryState: "Information",
                    validPct: 0,
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
                    let excelData = XLSX.utils.sheet_to_row_object_array(oSheet, { defval: "" });

                    if (!excelData || excelData.length === 0) {
                        MessageToast.show("File rỗng hoặc thiếu dữ liệu!");
                        this._refreshTable();
                        return;
                    }

                    const hasHeader = this._hasHeaderRow(excelData[0]);
                    this._hasHeaderRowFlag = hasHeader;
                    this._prepareMappingPreview(excelData[0], hasHeader);

                    if (hasHeader) {
                        excelData = excelData.slice(1);
                    }
                    this._rawExcelData = excelData;
                    const result = this._processRows(excelData);
                    this.dataUpload = result.data;
                    oModel.setProperty("/items", this.dataUpload);
                    this._applyUploadSummary(this.dataUpload);

                    const oSummary = oModel.getProperty("/summary");
                    this.byId(CHECK_BTN_ID).setEnabled(this.dataUpload.length > 0);
                    this.byId(POST_BTN_ID).setEnabled(oSummary.valid > 0 && oSummary.invalid === 0);

                    if (this.dataUpload.length === 0) {
                        MessageToast.show("Không có dòng dữ liệu hợp lệ!");
                        return;
                    }
                    if (oSummary.invalid > 0) {
                        MessageToast.show("File có " + oSummary.invalid + " dòng lỗi. Vui lòng sửa file và upload lại.");
                    } else {
                        MessageToast.show("Đọc thành công " + this.dataUpload.length + " dòng");
                    }
                } catch (err) {
                    MessageBox.error(err?.message || "Lỗi đọc file.");
                    this.getView().getModel(UPLOAD_MODEL).setProperty("/mapping", []);
                    this.getView().getModel(UPLOAD_MODEL).setProperty("/mappingActive", false);
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

                excelData.forEach((raw, index) => {
                    if (!raw) return;
                    const rowData = {
                        gr_number:       _key(raw, this._getMappingHeader("gr_number")),
                        document_date:   this._convDate(_key(raw, this._getMappingHeader("document_date"))),
                        movement_type:   _key(raw, this._getMappingHeader("movement_type")) || "101",
                        po_number:       _key(raw, this._getMappingHeader("po_number")),
                        po_item:         _key(raw, this._getMappingHeader("po_item")),
                        receive_qty:     _key(raw, this._getMappingHeader("receive_qty")),
                        unit:            _key(raw, this._getMappingHeader("unit")),
                        storage_location: _key(raw, this._getMappingHeader("storage_location")),
                    };

                    const aErrors = this._validateRow(rowData);
                    rowData.ValidationStatus = aErrors.length === 0 ? "S" : "E";
                    rowData.ValidationMessage = aErrors.join("; ");
                    rowData.RowIndex = index + (this._hasHeaderRowFlag ? 2 : 1);

                    if (!_key(raw, "po_number") && !_key(raw, "gr_number")) return;
                    aData.push(rowData);
                });
                return { data: aData };
            },

            _normalizeHeader(value) {
                return String(value || "").trim().toLowerCase().replace(/[_\s\-]+/g, " ");
            },

            _matchHeaderToField(header, field) {
                const normalized = this._normalizeHeader(header);
                if (!normalized) return false;
                const target = this._normalizeHeader(field.targetField);
                const label = this._normalizeHeader(field.label);

                if (normalized === target || normalized === label) {
                    return true;
                }
                if (target === "movement_type" && /movement|mvt|type/.test(normalized)) {
                    return true;
                }
                if (target === "receive_qty" && /qty|quantity|số lượng|receive/.test(normalized)) {
                    return true;
                }
                if (target === "storage_location" && /storage|location|sloc|kho/.test(normalized)) {
                    return true;
                }
                if (target === "po_item" && /po item|item/.test(normalized)) {
                    return true;
                }
                if (target === "po_number" && /po|purchase order/.test(normalized)) {
                    return true;
                }
                return false;
            },

            _hasHeaderRow(sampleRow) {
                if (!sampleRow || typeof sampleRow !== "object") {
                    return false;
                }
                return Object.keys(sampleRow).some((header) =>
                    GR_FIELD_MAPPING.some((field) => this._matchHeaderToField(header, field))
                );
            },

            _getMappingHeader(targetField) {
                const oModel = this.getView().getModel(UPLOAD_MODEL);
                const aMapping = oModel.getProperty("/mapping") || [];
                const oMap = aMapping.find((mapping) => mapping.targetField === targetField);
                return oMap?.sourceHeader || targetField;
            },

            _prepareMappingPreview(sampleRow, hasHeader) {
                const aHeaders = sampleRow && typeof sampleRow === "object"
                    ? Object.keys(sampleRow)
                    : [];
                const aSourceHeaders = [
                    { key: "", text: "--Chọn cột--" },
                    ...aHeaders.map((header) => ({ key: header, text: header }))
                ];
                const aMapping = GR_FIELD_MAPPING.map((field) => {
                    const sourceHeader = aHeaders.find((header) => this._matchHeaderToField(header, field)) || "";
                    return {
                        label: field.label,
                        targetField: field.targetField,
                        sourceHeader,
                        required: field.required ? "Yes" : "No",
                    };
                });
                const oModel = this.getView().getModel(UPLOAD_MODEL);
                oModel.setProperty("/mapping", aMapping);
                oModel.setProperty("/sourceHeaders", aSourceHeaders);
                oModel.setProperty("/mappingActive", false);
                oModel.setProperty("/hasHeaderRow", Boolean(hasHeader));
            },

            onToggleMappingPreview() {
                const oModel = this.getView().getModel(UPLOAD_MODEL);
                const bActive = oModel.getProperty("/mappingActive");
                oModel.setProperty("/mappingActive", !bActive);
            },

            _extractValue(raw, sourceHeader, targetField) {
                const val = String(raw[sourceHeader] || raw[sourceHeader?.toUpperCase?.()] || "").trim();
                if (val) return val;
                return String(raw[targetField] || raw[targetField?.toUpperCase?.()] || "").trim();
            },

            _reprocessRows() {
                if (!this._rawExcelData) return;
                const result = this._processRows(this._rawExcelData);
                this.dataUpload = result.data;
                const oModel = this.getView().getModel(UPLOAD_MODEL);
                oModel.setProperty("/items", this.dataUpload);
                this._applyUploadSummary(this.dataUpload);
                const oSummary = oModel.getProperty("/summary");
                this.byId(CHECK_BTN_ID).setEnabled(this.dataUpload.length > 0);
                this.byId(POST_BTN_ID).setEnabled(oSummary.valid > 0 && oSummary.invalid === 0);
            },

            onMappingChange(oEvent) {
                const sKey = oEvent.getParameter("selectedItem")?.getKey() || "";
                const oContext = oEvent.getSource().getBindingContext(UPLOAD_MODEL);
                if (oContext) {
                    oContext.setProperty("sourceHeader", sKey);
                    this._reprocessRows();
                }
            },

            _validateRow(rowData) {
                const aErrors = [];
                GR_FIELD_MAPPING.forEach((field) => {
                    if (field.required && !rowData[field.targetField]) {
                        aErrors.push(`${field.label} không được để trống`);
                    }
                });
                if (rowData.receive_qty && isNaN(Number(rowData.receive_qty))) {
                    aErrors.push("Receive Qty phải là số");
                }
                if (rowData.document_date && !/^\d{8}$/.test(rowData.document_date)) {
                    aErrors.push("Document Date phải đúng định dạng YYYYMMDD");
                }
                return aErrors;
            },

            _convDate(v) {
                if (!v && v !== 0) return "";
                const s = String(v).trim();
                const ddmmyyyy = /^(\d{2})[\\/](\d{2})[\\/](\d{4})$/;
                const yyyymmdd = /^(\d{4})[-\\/](\d{2})[-\\/](\d{2})$/;
                if (ddmmyyyy.test(s)) {
                    return s.replace(ddmmyyyy, "$3$2$1");
                }
                if (yyyymmdd.test(s)) {
                    return s.replace(yyyymmdd, "$1$2$3");
                }
                if (!isNaN(Number(s)) && s.length <= 8) {
                    const n = Number(s);
                    if (n >= 19000101 && n <= 29991231) {
                        return String(n).padStart(8, "0");
                    }
                    if (n > 59) {
                        const date = new Date(Math.round((n - 25569) * 86400 * 1000));
                        const pad = (x) => String(x).padStart(2, "0");
                        return `${date.getFullYear()}${pad(date.getMonth() + 1)}${pad(date.getDate())}`;
                    }
                }
                return s;
            },

            _applyUploadSummary(aData) {
                const oModel = this.getView().getModel(UPLOAD_MODEL);
                const iTotal = aData.length;
                const iValid = aData.filter((r) => r.ValidationStatus === "S").length;
                const iInvalid = iTotal - iValid;
                const iPct = iTotal === 0 ? 0 : Math.round((iValid / iTotal) * 100);
                const sState = iTotal === 0 ? "Information"
                    : iInvalid === 0 ? "Success"
                    : iValid === 0 ? "Error"
                    : "Warning";

                oModel.setProperty("/summary", {
                    total: iTotal,
                    valid: iValid,
                    invalid: iInvalid,
                    successRate: iPct,
                    errorRate: 100 - iPct,
                });
                oModel.setProperty("/summaryText",
                    iTotal === 0
                        ? "Chưa có dữ liệu"
                        : "Tổng: " + iTotal + " dòng · Hợp lệ: " + iValid + " · Lỗi: " + iInvalid + " (" + iPct + "%)"
                );
                oModel.setProperty("/summaryState", sState);
                oModel.setProperty("/validPct", iPct);
            },

            formatIsError: function (sStatus) {
                return String(sStatus) === "E";
            },

            shortBatchId: function (sBatchId) {
                if (!sBatchId) return "";
                const s = String(sBatchId);
                return s.length <= 8 ? s : s.substring(0, 8) + "…";
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
                    const sStatusText = oResult.status === "S" ? "Hợp lệ" :
                                        oResult.status === "E" ? "Lỗi" : "Đang xử lý";
                    const sStatusState = oResult.status === "S" ? "Success" :
                                         oResult.status === "E" ? "Error" : "Warning";
                    oResultModel.setData({
                        ...oResult,
                        statusText: sStatusText,
                        statusState: sStatusState,
                    });
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

            onApplyHistoryFilter() {
                const aFilters = [];
                const sStatus = this.byId("grFilterStatus")?.getSelectedKey();
                const sSearch = this.byId("grFilterSearch")?.getValue();
                const sBatch  = this.byId("grFilterBatch")?.getValue();

                if (sStatus) aFilters.push(new Filter("Status", FilterOperator.EQ, sStatus));
                if (sSearch) aFilters.push(new Filter("GrNumber", FilterOperator.Contains, sSearch));
                if (sBatch)  aFilters.push(new Filter("BatchId", FilterOperator.Contains, sBatch));

                const oBinding = this.byId("grHistoryTable")?.getBinding("items");
                if (oBinding) oBinding.filter(aFilters);
            },
            onResetHistoryFilter() {
                this.byId("grFilterStatus").setSelectedKey("");
                this.byId("grFilterSearch").setValue("");
                this.byId("grFilterBatch").setValue("");
                this.onApplyHistoryFilter();
            },

                        onRetrySelected: function () {
                const aItems = this.byId("grHistoryTable").getSelectedItems();
                const aContexts = aItems
                    .map((it) => it.getBindingContext("gr"))
                    .filter((ctx) => ctx && ctx.getProperty("Status") === "E");

                if (aContexts.length === 0) {
                    MessageToast.show("Chưa chọn dòng Lỗi nào (chỉ retry được dòng Status = Lỗi).");
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

            // ── Xem chi tiết item (dùng chung cho tab Chờ xử lý và Lịch sử) ──
            onShowItems: async function (oEvent) {
                const oContext = oEvent.getParameter
                    ? oEvent.getParameter("listItem")?.getBindingContext("gr")
                    : null;
                const oCtx = oContext || oEvent.getSource().getBindingContext("gr");
                if (!oCtx) return;
                const sGrNumber = oCtx.getProperty("GrNumber");

                if (!this._oItemDialog) {
                    this._oItemDialog = await this.loadFragment({ name: "zfipkzup.view.fragment.GrItemDetail" });
                }
                this._oItemDialog.setModel(new JSONModel({ grNumber: sGrNumber }), "itemDialog");

                // Access table from dialog's content and apply filter for the selected GR
                const oTable = this._oItemDialog.getContent()[0]; // Table is first content element
                const oItemBinding = oTable?.getBinding("items");
                if (oItemBinding) {
                    oItemBinding.filter([new Filter("GrNumber", FilterOperator.EQ, sGrNumber)]);
                    const aContexts = await oItemBinding.requestContexts(0, 500);
                    const aItems = aContexts.map((oCtx) => oCtx.getObject()).filter(Boolean);
                    const iTotal = aItems.length;
                    const iInvalid = aItems.filter((item) => item.Status === "E").length;
                    const aErrors = aItems
                        .filter((item) => item.Status === "E")
                        .map((item) => ({
                            title: "Item " + item.Item,
                            description: item.Message || "Lỗi chưa có chi tiết",
                        }));

                    const oItemDialogModel = new JSONModel({
                        grNumber: sGrNumber,
                        totalItems: iTotal,
                        validItems: iTotal - iInvalid,
                        invalidItems: iInvalid,
                        summaryState: iInvalid > 0 ? "Error" : "Success",
                        hasErrors: iInvalid > 0,
                        errors: aErrors,
                    });
                    this._oItemDialog.setModel(oItemDialogModel, "itemDialog");
                }
                this._oItemDialog.open();
            },
            onCloseItemDialog() {
                this._oItemDialog?.close();
            },

            onHistoryTableUpdateFinished() {
                const oTable = this.byId("grHistoryTable");
                if (!oTable) return;
                const aContexts = oTable.getBinding("items")?.getContexts(0, 1000) || [];
                const bHasCreatedBy = aContexts.some((oCtx) => {
                    const v = oCtx.getProperty("CreatedBy");
                    return v !== undefined && String(v).trim() !== "";
                });
                const bHasCreatedAt = aContexts.some((oCtx) => {
                    const v = oCtx.getProperty("CreatedAt");
                    return v !== undefined && String(v).trim() !== "";
                });
                const oModel = this.getView().getModel(UPLOAD_MODEL);
                oModel.setProperty("/showCreatedByColumn", bHasCreatedBy);
                oModel.setProperty("/showCreatedAtColumn", bHasCreatedAt);
            },

            onCopyBatchId: function (oEvent) {
                const sBatchId = this.getView().getModel("gr").getProperty(oEvent.getSource().getBindingContext("gr").getPath() + "/BatchId");
                if (!sBatchId) {
                    MessageToast.show("Không có Batch ID để copy.");
                    return;
                }
                if (navigator.clipboard && navigator.clipboard.writeText) {
                    navigator.clipboard.writeText(sBatchId)
                        .then(() => MessageToast.show("Đã copy Batch ID"))
                        .catch(() => MessageToast.show("Copy Batch ID không thành công"));
                } else {
                    MessageToast.show("Trình duyệt không hỗ trợ copy tự động.");
                }
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
                const oModel = this.getView().getModel(UPLOAD_MODEL);
                oModel.setProperty("/items", []);
                oModel.setProperty("/mapping", []);
                oModel.setProperty("/mappingActive", false);
                oModel.setProperty("/hasHeaderRow", false);
                oModel.setProperty("/summary", {
                    total: 0,
                    valid: 0,
                    invalid: 0,
                    successRate: 0,
                    errorRate: 0,
                });
                this.dataUpload = [];
                this._sFileName = "";
                this._hasHeaderRowFlag = false;
                this.byId(POST_BTN_ID).setEnabled(false);
                this.byId(CHECK_BTN_ID).setEnabled(false);
                this.byId("grFileUploader")?.clear();
            },
        });
    }
);
