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
                this.getView().setModel(new JSONModel({ items: [], selectedCount: 0 }), "poLookup");
                this.onApplyHistoryFilter();
                this._startPolling();
                // this._loadAuth();
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

            onSearchOpenPO: async function () {
                const sPo = this.byId("poLookupPoNumber").getValue().trim();
                const sPlant = this.byId("poLookupPlant").getValue().trim();
                const sMaterial = this.byId("poLookupMaterial").getValue().trim();
                const sVendor = this.byId("poLookupVendor").getValue().trim();
                const sPurchGroup = this.byId("poLookupPurchGroup").getValue().trim();
                const sPurchOrg = this.byId("poLookupPurchOrg").getValue().trim();
                const oDateFrom = this.byId("poLookupDateFrom").getDateValue();
                const oDateTo = this.byId("poLookupDateTo").getDateValue();
                const bHideRecv = this.byId("poLookupHideReceived").getSelected();

                const pad = (n) => String(n).padStart(2, "0");
                const fmt = (d) => d ? `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}` : "";

                this._showBusy();
                try {
                    const aRows = await ApiService.searchOpenPO(this._getGrModel(), {
                        poNumber: sPo,
                        plant: sPlant,
                        material: sMaterial,
                        vendor: sVendor,
                        purchasingGroup: sPurchGroup,
                        purchasingOrg: sPurchOrg,
                        dateFrom: fmt(oDateFrom),
                        dateTo: fmt(oDateTo),
                        hideReceived: bHideRecv,
                    });
                    this.getView().getModel("poLookup").setProperty("/items", aRows);
                    this.getView().getModel("poLookup").setProperty("/selectedCount", 0);
                    this.byId("poLookupTable").removeSelections();
                } catch (e) {
                    MessageBox.error(e.message || "Không tra được PO");
                } finally {
                    this._closeBusy();
                }
            },


            onPoLookupSelectionChange: function () {
                const iCount = this.byId("poLookupTable").getSelectedItems().length;
                this.getView().getModel("poLookup").setProperty("/selectedCount", iCount);
            },

            onDownloadPrefilledTemplate: function () {
                const aSelected = this.byId("poLookupTable").getSelectedItems()
                    .map((oItem) => oItem.getBindingContext("poLookup").getObject());
                if (!aSelected.length) return;

                sap.ui.require(["zfipkzup/controller/helper/GrExcelTemplate"], function (GrExcelTemplate) {
                    GrExcelTemplate.downloadPrefilled(aSelected);
                });
            },

            onOpenCreateFromPO: async function (oEvent) {
                const oRow = oEvent.getSource().getBindingContext("poLookup").getObject();

                if (!this._oCreatePODialog) {
                    this._oCreatePODialog = await this.loadFragment({ name: "zfipkzup.view.fragment.CreateFromPO" });
                }

                this.getView().setModel(new JSONModel({
                    poNumber: oRow.PurchaseOrder,
                    poItem: oRow.PurchaseOrderItem,
                    material: oRow.Material,
                    shortText: oRow.ShortText,
                    orderUnit: oRow.OrderUnit,
                    openQuantity: oRow.OpenQuantity,
                    grNumber: "",
                    receiveQty: null,
                    storageLocation: oRow.StorageLocation,
                    batch: "",
                }), "createPO");

                this.byId("cfpDocDate").setDateValue(new Date());
                this._oCreatePODialog.open();
            },

            onCancelCreateFromPO: function () {
                this._oCreatePODialog?.close();
            },

            onConfirmCreateFromPO: async function () {
                const oModel = this.getView().getModel("createPO");
                const d = oModel.getData();
                const oDate = this.byId("cfpDocDate").getDateValue();

                if (!d.grNumber || !d.receiveQty || !d.storageLocation || !oDate) {
                    MessageToast.show("Nhập đủ GR Number / Document Date / Receive Qty / Storage Location");
                    return;
                }

                const pad = (n) => String(n).padStart(2, "0");
                const sDocDate = oDate.getFullYear() + "-" + pad(oDate.getMonth() + 1) + "-" + pad(oDate.getDate());

                this._showBusy();
                try {
                    const oResult = await ApiService.callActionCreateFromPO(this._getGrModel(), {
                        po_number: d.poNumber,
                        po_item: d.poItem,
                        gr_number: d.grNumber,
                        document_date: sDocDate,
                        receive_qty: String(d.receiveQty),
                        unit: d.orderUnit,
                        storage_location: d.storageLocation,
                        batch: d.batch,
                    });
                    console.log("GR upload result:", JSON.stringify(oResult));

                    if (oResult.status === "E") {
                        MessageBox.error(oResult.message || "Không tạo được GR nháp");
                        return;
                    }

                    MessageToast.show("Đã tạo GR " + d.grNumber + " (nháp) — sang tab Chờ xử lý để Post");
                    this._oCreatePODialog.close();
                    this.onApplyHistoryFilter();
                    this.byId("idIconTabBarGr").setSelectedKey("pending");
                } catch (e) {
                    MessageBox.error(e.message || "Lỗi không xác định");
                } finally {
                    this._closeBusy();
                }
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
    this.getView().getModel("grResult").setData({});
    this.byId("grResultPanel").setVisible(false);
    this._showBusy();

                try {
                    await ApiService.checkFileExistsGR(this._getGrModel(), this._sFileName);

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

                    // Chặn trùng GR Number ngay lúc chọn file: file mới nhưng GR Number
                    // đã tồn tại (kể cả bản lỗi E ở Lịch sử) thì không cho Check/Post.
                    const aGrNumbers = aPrepared
                        .map((o) => String(o.r.gr_number || "").trim().toUpperCase())
                        .filter(Boolean);
                    const aExistingGr = await ApiService.checkGrNumbersExist(this._getGrModel(), aGrNumbers);
                    if (aExistingGr.length > 0) {
                        const mStatusText = {
                            S: "đã post thành công", P: "đang chờ job xử lý",
                            R: "đang là nháp ở tab Chờ xử lý", E: "đang LỖI ở tab Lịch sử"
                        };
                        ErrorDialog.handleErrorDialog(
                            aExistingGr.map((o) => ({
                                type: "Error",
                                title: `GR ${o.grNumber} đã tồn tại`,
                                description: `GR Number này ${mStatusText[o.status] || "đã tồn tại"} — ` +
                                    `đổi GR Number khác trong file, hoặc xử lý bản cũ ở tab Chờ xử lý / Lịch sử.`,
                            })),
                            this
                        );
                        this.byId(POST_BTN_ID).setEnabled(false);
                        this.byId(CHECK_BTN_ID).setEnabled(false);
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

                    // OData action trả PascalCase (Status/Message/TotalCount) — chuẩn hóa
                    // về 1 biến, tránh đọc oResult.status (lowercase) ra undefined → toast xanh nhầm
                    const sStatus = oResult.status ?? oResult.Status ?? "";
                    const sMessage = oResult.message ?? oResult.Message ?? "";
                    const iTotal = oResult.total_count ?? oResult.TotalCount ?? 0;

                    const oResultModel = this.getView().getModel("grResult");
                   if (bTestMode) {
    oResultModel.setData({
        ...oResult,
        statusText: sStatus === "S" ? "Hợp lệ" :
            sStatus === "E" ? "Lỗi" : "Đang xử lý",
        statusState: sStatus === "S" ? "Success" :
            sStatus === "E" ? "Error" : "Warning",
    });

    // Panel xanh ở trên chỉ là check cú pháp Excel — nếu SAP đã từ chối thật
    // (period, PO, qty...) thì đây mới là verdict cuối, phải đồng bộ lại,
    // không để xanh "100% hợp lệ" đứng cạnh đỏ "Lỗi" của cùng 1 lần Check.
    if (sStatus === "E") {
        const oGrModel = this.getView().getModel(UPLOAD_MODEL);
        oGrModel.setProperty("/summaryText", "SAP từ chối: " + sMessage);
        oGrModel.setProperty("/summaryState", "Error");
    }
}

                    this.byId("grResultPanel").setVisible(true);

                   // Tách lỗi theo từng dòng để hiện dialog chi tiết
                    let aErrItems = [];
                    try {
                        const aAll = JSON.parse(oResult.items_json ?? oResult.ItemsJson ?? "[]") || [];
                        aErrItems = aAll.filter((it) => (it.status ?? it.Status) === "E");
                    } catch (e) { /* rỗng → fallback message */ }

                    const oBundle = this.getView().getModel("i18n").getResourceBundle();

                    if (sStatus === "E") {
                        if (aErrItems.length) {
                            ErrorDialog.handleErrorDialog(
                                aErrItems.map((it) => ({
                                    type: "Error",
                                    title: `GR ${it.grNumber ?? it.GrNumber} — Item ${it.item ?? it.Item}` +
                                           ` (PO ${it.poNumber ?? it.PoNumber}/${it.poItem ?? it.PoItem})`,
                                    description: it.message ?? it.Message,
                                })),
                                this
                            );
                        } else {
                            MessageBox.error(sMessage || "Check thất bại.");
                        }
                    } else if (bTestMode) {
                        MessageToast.show(oBundle.getText("CHECK_COMPLETE"));
                    } else {
                        MessageToast.show(oBundle.getText("SENT_PROCESSING", [iTotal, "GR"]));
                    }

                    if (bTestMode) {
                        this.onRefreshPending();
                    } else {
                        this._refreshPendingAndHistory();
                    }

                } catch (err) {
                    MessageBox.error(this._extractError(err, "Lỗi khi gọi SAP."));
                } finally {
                    this._closeBusy();
                }
            },

            /**
 * Tra open PO qua entity đọc-thôi POItem (Check Open PO). Không có PO Number
 * lẫn Plant thì không query — tránh kéo toàn bộ PO đang mở trong hệ thống.
 */
            searchOpenPO: async function (oModel, sPoNumber, sPlant) {
                const aFilters = [];
                if (sPoNumber) {
                    aFilters.push(new Filter("PurchaseOrder", FilterOperator.EQ, sPoNumber));
                }
                if (sPlant) {
                    aFilters.push(new Filter("Plant", FilterOperator.EQ, sPlant));
                }
                const oBinding = oModel.bindList("/POItem", undefined, [], aFilters, {
                    $orderby: "PurchaseOrder,PurchaseOrderItem",
                });
                const aContexts = await oBinding.requestContexts(0, 200);
                return aContexts.map((oContext) => oContext.getObject());
            },

            /**
             * Tạo 1 dòng GR nháp (status R) trực tiếp từ 1 PO item đã chọn ở panel
             * Check Open PO — dùng action createFromPO, KHÔNG đụng uploadExcel.
             */
            callActionCreateFromPO: async function (oModel, oParam) {
                const ACTION_FQN_CREATE_PO =
                    "com.sap.gateway.srvd.zmm_ui_pogr_o4.v0001.createFromPO";

                const oOperation = oModel.bindContext(
                    "/GrUpload/" + ACTION_FQN_CREATE_PO + "(...)"
                );
                oOperation.setParameter("po_number", oParam.po_number);
                oOperation.setParameter("po_item", oParam.po_item);
                oOperation.setParameter("gr_number", oParam.gr_number);
                oOperation.setParameter("document_date", oParam.document_date);
                oOperation.setParameter("receive_qty", oParam.receive_qty);
                oOperation.setParameter("unit", oParam.unit);
                oOperation.setParameter("storage_location", oParam.storage_location);
                oOperation.setParameter("batch", oParam.batch || "");
                oOperation.setParameter("user_email", await this.getUserEmail());
                await oOperation.execute();

                return oOperation.getBoundContext().getObject() || {};
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
                        // this.byId("fileUploader")?.setEnabled(false);   // ← THÊM: khóa nút chọn file
                        // this._bNoAuth = true;
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
