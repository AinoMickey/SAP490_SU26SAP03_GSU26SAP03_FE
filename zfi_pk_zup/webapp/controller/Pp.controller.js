sap.ui.define(
  [
    "sap/ui/core/mvc/Controller",
    "sap/m/MessageToast",
    "sap/ui/model/json/JSONModel",
    "sap/m/MessageBox",
    "sap/ui/core/library",
    "./helper/PpExcelTemplate",
    "./helper/ExcelParser",
    "./helper/ApiService",
    "./helper/ResultsDialog",
    "./xlsx/xlsx.bundle",
    "./helper/HistoryDetailDialog",
  ],
  function (
    Controller,
    MessageToast,
    JSONModel,
    MessageBox,
    coreLibrary,
    PpExcelTemplate,
    ExcelParser,
    ApiService,
    ResultsDialog,
    xlsxBundle, // SỬA: trước đây THIẾU tham số này -> 11 dependency / 10 tham số,
    HistoryDetailDialog, //  làm HistoryDetailDialog thực chất trỏ vào module xlsx.
  ) {
    "use strict";

    const ValueState = coreLibrary.ValueState;
    const UPLOAD_MODEL = "ppUploadModel";
    const TABLE_ID = "ppTableUpload";
    const POST_BTN_ID = "ppPostButton";
    const CHECK_BTN_ID = "ppCheckButton";

    const KEY_ALIASES = {
      startbasicdates: "datestart",
      endbasicdates: "dateend",
      salesorder: "saleorder",
      salesorderitem: "saleorderitem",
      note: "longtext",
    };

    return Controller.extend("zfipkzup.controller.Pp", {
      dataUpload: [],

      onInit() {
        this.getView().setModel(new JSONModel({ items: [] }), UPLOAD_MODEL);
      },

      onExit() {
        this._destroyBusy();
      },

      _getPpModel() {
        // Model OData V4 trỏ service zpp_ui_zuplsx_o4 (manifest model "pp")
        return this.getView().getModel("pp");
      },

      _getBundle() {
        const oModel = this.getView().getModel("i18n");
        return oModel ? oModel.getResourceBundle() : null;
      },

      /**
       * getText trả về CHÍNH KEY khi thiếu entry trong i18n.properties,
       * nên `getText(k) || fallback` không bao giờ chạy tới fallback mà
       * đẩy thẳng tên key ra màn hình. Hàm này so với key nên fallback
       * mới thật sự có tác dụng.
       */
      _t(sKey, aArgs, sFallback) {
        const oBundle = this._getBundle();
        if (!oBundle) return sFallback;
        const s = oBundle.getText(sKey, aArgs);
        return !s || s === sKey ? sFallback : s;
      },

      // ── Navigation ──

      onNavBack() {
        this.getOwnerComponent().getRouter().navTo("RouteMain");
      },

      // ── Template ──

      onDownloadTemplate() {
        PpExcelTemplate.download();
        MessageToast.show(
          this._t("TEMPLATE_LOADING", null, "Đang tải template..."),
        );
      },

      // ── Upload / Parse ──

      onFileChange: async function (oEvent) {
        const oFile =
          oEvent.getParameter("files") && oEvent.getParameter("files")[0];
        const oModel = this.getView().getModel(UPLOAD_MODEL);

        if (!oFile) {
          this._refreshTable();
          return;
        }

        this.dataUpload = [];
        oModel.setProperty("/items", []);
        this._showBusy();

        try {
          // check trùng file
          await ApiService.checkFileExistsPP(this._getPpModel(), oFile.name);

          const fileContent = await ExcelParser.readFile(oFile);
          const workbook = XLSX.read(fileContent, { type: "binary" });
          const oSheet =
            workbook.Sheets["Data"] || workbook.Sheets[workbook.SheetNames[0]];
          const excelData = XLSX.utils.sheet_to_row_object_array(oSheet);

          // ══════════ 1: nạp validator + check template ══════════
          const UploadValidator = await new Promise((res) =>
            sap.ui.require(["zfipkzup/controller/helper/UploadValidator"], res),
          );
          const ErrorDialog = await new Promise((res) =>
            sap.ui.require(["zfipkzup/controller/helper/ErrorDialog"], res),
          );

          const aHeader =
            XLSX.utils.sheet_to_json(oSheet, { header: 1 })[0] || [];
          const oTpl = UploadValidator.checkTemplate(aHeader, "PP");
          if (!oTpl.ok) {
            ErrorDialog.handleErrorDialog(
              [
                {
                  type: "Error",
                  title: this._t("EXCEL_TEMPLATE_INVALID", ["PP"],
                    "File không đúng template PP"),
                  description: this._t("TEMPLATE_MISSING_COLS",
                    [oTpl.missing.join(", ")],
                    "Thiếu cột: " + oTpl.missing.join(", ") +
                    ". Vui lòng tải Template mới và điền lại."),
                },
              ],
              this,
            );
            this._refreshTable();
            return;
          }

          // ══════════ 2: bóc label an toàn (THAY delete excelData[0]) ══════════
          const oStrip = UploadValidator.stripLabelRow(excelData);
          const aPrepared = UploadValidator.prepareRows(
            oStrip.rows,
            oStrip.startRow,
          );

          // check file rỗng: user xóa label row thì file 1 dòng data vẫn hợp lệ
          if (aPrepared.length === 0) {
            MessageToast.show(
              this._t("UPLOAD_NO_DATA", null, "File không có dòng dữ liệu!"),
            );
            this._refreshTable();
            return;
          }

          // ══════════ 3: chặn syntax ══════════
          const aSynErrors = UploadValidator.validatePP(aPrepared);
          if (aSynErrors.length > 0) {
            ErrorDialog.handleErrorDialog(
              aSynErrors.map((e) => ({
                type: "Error",
                title: this._t("EXCEL_ROW_COL", [e.excelRow, e.field],
                  "Dòng Excel " + e.excelRow + " — cột " + e.field),
                description: e.message,
              })),
              this,
            );
            this._refreshTable();
            return;
          }

          // ══════════ TỪ ĐÂY GIỮ NGUYÊN FLOW CŨ, chỉ đổi nguồn data ══════════
          const result = this._processRows([null].concat(oStrip.rows));
          this.dataUpload = result.data;
          oModel.setProperty("/items", this.dataUpload);
          this._autoSizeColumns(this.dataUpload);

          if (result.invalidRows.length > 0) {
            MessageBox.error(
              this._t("PP_INVALID_DATE_ROWS",
                [result.invalidRows.length, result.invalidRows.join(", ")],
                "Sai định dạng ngày (DD/MM/YYYY) ở " +
                result.invalidRows.length + " dòng: " +
                result.invalidRows.join(", ")),
            );
            this.byId(POST_BTN_ID).setEnabled(false);
            this.byId(CHECK_BTN_ID).setEnabled(false);
          } else {
            MessageToast.show(
              this._t("UPLOAD_SUCCESS", [this.dataUpload.length],
                "Tải lên thành công " + this.dataUpload.length + " dòng"),
            );
            this.byId(POST_BTN_ID).setEnabled(true);
            this.byId(CHECK_BTN_ID).setEnabled(true);
          }
        } catch (error) {
          MessageBox.error(
            error?.message ||
              this._t("UPLOAD_CANCELLED", null,
                "Tải lên bị hủy hoặc thất bại."),
          );
          this.byId(POST_BTN_ID).setEnabled(false);
          this.byId(CHECK_BTN_ID).setEnabled(false);
        } finally {
          this._closeBusy();
        }
      },

      onHistoryItemPress: function (oEvent) {
        const oCtx = oEvent.getSource().getBindingContext("pp");
        if (!oCtx) return;
        const o = oCtx.getObject();
        const oRptModel = this.getOwnerComponent().getModel("rpt");
        const oView = this.getView();
        const that = this;

        sap.ui.require(
          [
            "zfipkzup/controller/helper/HistoryDetailDialog",
            "sap/ui/model/Filter",
            "sap/ui/model/FilterOperator",
          ],
          async function (HistoryDetailDialog, Filter, FilterOperator) {
            // Lấy bản ghi đầy đủ (kèm OrderStatus/ReleaseDate) từ RPT
            let oFull = o;
            if (oRptModel && o.ProductionOrder) {
              try {
                const oBinding = oRptModel.bindList(
                  "/PPUploadReport",
                  null,
                  [],
                  [
                    new Filter(
                      "ProductionOrder",
                      FilterOperator.EQ,
                      o.ProductionOrder,
                    ),
                  ],
                );
                const aCtx = await oBinding.requestContexts(0, 1);
                if (aCtx.length) oFull = aCtx[0].getObject();
              } catch (e) {
                /* fallback dữ liệu từ bảng */
              }
            }

            HistoryDetailDialog.open({
              view: oView,
              title: that._t("PP_DETAIL_TITLE", [oFull.ProductionOrder || ""],
                "Chi tiết lệnh " + (oFull.ProductionOrder || "")),
              record: oFull,
              fields: [
                { key: "IdDoc", label: that._t("LBL_ID", null, "ID") },
                { key: "ProductionOrder", label: that._t("LBL_PRODUCTION_ORDER", null, "Production Order") },
                { key: "OrderStatus", label: that._t("LBL_ORDER_STATUS", null, "Trạng thái") },
                { key: "ReleaseDate", label: that._t("LBL_RELEASE_DATE", null, "Release Date"), type: "date" },
                { key: "ProductionPlant", label: that._t("LBL_PLANT", null, "Plant") },
                { key: "Material", label: that._t("LBL_MATERIAL", null, "Material") },
                { key: "ProductionOrderType", label: that._t("LBL_ORDER_TYPE", null, "Order Type") },
                { key: "ProductionVersion", label: that._t("LBL_PRODUCTION_VERSION", null, "Production Version") },
                { key: "TotalQty", label: that._t("LBL_TOTAL_QTY", null, "Total Qty"), type: "num" },
                { key: "BaseUnit", label: that._t("LBL_UNIT", null, "Unit") },
                { key: "StartDate", label: that._t("LBL_START_DATE", null, "Start Date"), type: "date" },
                { key: "EndDate", label: that._t("LBL_END_DATE", null, "End Date"), type: "date" },
                { key: "LeadTimeDays", label: that._t("LBL_LEAD_TIME", null, "Lead time (ngày)"), type: "num" },
                { key: "SalesOrder", label: that._t("LBL_SALES_ORDER", null, "Sales Order") },
                { key: "SalesOrderItem", label: that._t("LBL_SO_ITEM", null, "SO Item") },
                { key: "Filename", label: that._t("LBL_FILENAME", null, "Filename") },
                { key: "PstDate", label: that._t("LBL_PST_DATE", null, "Ngày post"), type: "date" },
                { key: "PstUser", label: that._t("LBL_PST_USER", null, "Người post") },
              ],
              copy: {
                label: that._t("COPY_ORDER_NO", null, "Copy số lệnh"),
                value: oFull.ProductionOrder,
              },
            });
          },
        );
      },

      onDownloadHistoryFile: function (oEvent) {
        const oCtx = oEvent.getSource().getBindingContext("pp");
        if (!oCtx) return;

        const sFilename = oCtx.getObject().Filename;
        const oPpModel = this._getPpModel();
        const that = this;

        if (!sFilename) {
          MessageBox.warning(
            this._t("PP_NO_FILENAME", null,
              "Dòng này không có Filename trong log (lệnh tạo trước khi log lưu filename), không dựng lại file được."),
          );
          return;
        }

        this._showBusy();

        sap.ui.require(
          ["zfipkzup/controller/helper/HistoryFileExport"],
          async function (HistoryFileExport) {
            try {
              const oRes = await HistoryFileExport.exportPp(
                oPpModel,
                sFilename,
                { withRefs: true },
              );
              MessageToast.show(
                that._t("PP_CREATE_HISTORY_FILE", [oRes.file, oRes.rows],
                  "Đã tạo " + oRes.file + " (" + oRes.rows +
                  " lệnh). Cột longtext không có trong log."),
              );
            } catch (e) {
              MessageBox.error(
                that._t("ERROR_BUILD_FILE", [e.message || e],
                  "Không dựng được file: " + (e.message || e)),
              );
            } finally {
              that._closeBusy();
            }
          },
          function (oErr) {
            that._closeBusy();
            MessageBox.error(
              that._t("ERROR_LOAD_MODULE",
                ["HistoryFileExport.js", oErr.message || oErr],
                "Không nạp được HistoryFileExport.js: " + (oErr.message || oErr)),
            );
          },
        );
      },

      /**
       * Chuẩn hóa + validate dữ liệu Excel.
       * - Key lowercase + alias từ template cloud cũ
       * - Bỏ dòng trống
       * - Ngày nhận DD/MM/YYYY hoặc DD.MM.YYYY, chuẩn hóa về DD/MM/YYYY
       *   (validator backend conv_date xử lý DD/MM/YYYY)
       */
      _processRows(excelData) {
        const aData = [];
        const aInvalidRows = [];
        const iStartRow = 3; // data bắt đầu từ row 3 trong Excel
        const sReady = this._t("PP_ROW_READY", null, "Ready");

        const normalizeKeys = (row) => {
          const out = {};
          Object.keys(row || {}).forEach((k) => {
            let sKey = k.toLowerCase().trim();
            if (KEY_ALIASES[sKey]) sKey = KEY_ALIASES[sKey];
            out[sKey] = row[k];
          });
          return out;
        };

        const normDate = (v) => {
          // Trả về { value: 'DD/MM/YYYY', valid: bool }
          if (v === undefined || v === null) return { value: "", valid: false };
          const s = String(v).trim().replace(/\./g, "/");
          const m = /^(0[1-9]|[12][0-9]|3[01])\/(0[1-9]|1[0-2])\/(\d{4})$/.exec(
            s,
          );
          if (!m) return { value: s, valid: false };
          const dd = Number(m[1]),
            mm = Number(m[2]),
            yyyy = Number(m[3]);
          const d = new Date(yyyy, mm - 1, dd);
          const bValid =
            d.getFullYear() === yyyy &&
            d.getMonth() === mm - 1 &&
            d.getDate() === dd;
          return { value: s, valid: bValid };
        };

        let iRowIndex = 0;
        excelData.forEach((raw, idx) => {
          if (!raw) return;
          const r = normalizeKeys(raw);

          // Dòng trống: không có cả material/plant/ordertype/totalqty
          if (
            !r.material &&
            !r.productionplant &&
            !r.ordertype &&
            !r.totalqty
          ) {
            return;
          }

          iRowIndex += 1;
          const iExcelRow = iStartRow + idx - 1; // số dòng thật trong Excel để báo lỗi
          const oStart = normDate(r.datestart);
          const oEnd = normDate(r.dateend);

          if (!oStart.valid || !oEnd.valid) {
            aInvalidRows.push(iExcelRow);
          }

          aData.push({
            clientRowId: String(iRowIndex),
            id_doc: String(r.id_doc || iRowIndex),
            ordertype: String(r.ordertype || ""),
            productionplant: String(r.productionplant || ""),
            material: String(r.material || ""),
            productionversion: String(r.productionversion || ""),
            totalqty: String(r.totalqty || ""),
            baseunit: String(r.baseunit || ""),
            datestart: oStart.value,
            dateend: oEnd.value,
            saleorder: String(r.saleorder || ""),
            saleorderitem: String(r.saleorderitem || ""),
            longtext: String(r.longtext || ""),

            // Kết quả post (đổ từ response theo clientRowId)
            productionorder: "",
            exceptionIcon: "",
            exceptionState: ValueState.None,
            exceptionText: sReady,

            // Trạng thái validate ngày phía client
            datestartState: oStart.valid ? ValueState.None : ValueState.Error,
            dateendState: oEnd.valid ? ValueState.None : ValueState.Error,
            datestartIcon: oStart.valid ? "" : "sap-icon://error",
            dateendIcon: oEnd.valid ? "" : "sap-icon://error",
          });
        });

        return { data: aData, invalidRows: aInvalidRows };
      },

      // ── Check / Post ──

      onCheck() {
        // Check: validate toàn bộ dòng (testmode X, backend không tạo lệnh)
        this._callUpload("X", this.dataUpload);
      },

      onPost: async function () {
        const bHasPosted = this.dataUpload.some(
          (item) => item.exceptionState === ValueState.Success,
        );
        const aPending = this.dataUpload.filter(
          (item) => item.exceptionState !== ValueState.Success,
        );
        if (aPending.length === 0) {
          MessageToast.show(
            this._t("ALL_POST_SUCCESS", null,
              "Tất cả các dòng đã post thành công."),
          );
          return;
        }
        try {
          // Chỉ check trùng file ở lần Post đầu. Nếu đã có dòng Success
          // (post một phần, đang re-post dòng lỗi) thì filename đã nằm
          // trong log do chính phiên này ghi -> bỏ check, không thì tự khóa mình.
          if (!bHasPosted) {
            await ApiService.checkFileExistsPP(
              this._getPpModel(),
              this.getCurrentFileName(),
            );
          }
          this._callUpload("", aPending);
        } catch (error) {
          MessageBox.error(error.message);
        }
      },

      _callUpload: async function (sTestMode, aItems) {
        this._showBusy();
        try {
          // Prevent double-submit
          this.byId(POST_BTN_ID)?.setEnabled(false);
          this.byId(CHECK_BTN_ID)?.setEnabled(false);

          const aRows = ApiService.buildPpRows(aItems);
          const oResponse = await ApiService.callActionUploadPP(
            this._getPpModel(),
            this.getCurrentFileName(),
            sTestMode,
            aRows,
          );

          this._applyResults(oResponse.results, sTestMode);

          // Truyền thêm results để dải tổng hợp liệt kê được số lệnh sản xuất
          ResultsDialog.show(
            this,
            oResponse.messages,
            sTestMode,
            POST_BTN_ID,
            oResponse.results,
          );

          // Post thật xong thì refresh tab Lịch sử
          if (sTestMode !== "X") {
            this.onRefreshHistory();
          }
        } catch (error) {
          // Lỗi gọi service thì mở lại nút để thử lại
          const bHasData = this.dataUpload.length > 0;
          this.byId(CHECK_BTN_ID)?.setEnabled(bHasData);
          this.byId(POST_BTN_ID)?.setEnabled(bHasData);
          MessageBox.error(error.message || JSON.stringify(error));
        } finally {
          this._closeBusy();
        }
      },

      /** Đổ kết quả từng dòng từ response vào cột Status theo clientRowId */
      _applyResults(aResults, sTestMode) {
        const oModel = this.getView().getModel(UPLOAD_MODEL);
        const mRows = new Map(
          this.dataUpload.map((item) => [item.clientRowId, item]),
        );

        (aResults || []).forEach((res) => {
          const oRow = mRows.get(String(res.ClientRowId));
          if (!oRow) return;

          const bSuccess = res.Type === "Success";
          oRow.exceptionText = res.Message;

          if (sTestMode === "X") {
            // Check: chỉ hiển thị kết quả validate, không khóa dòng
            oRow.exceptionIcon = bSuccess
              ? "sap-icon://sys-enter"
              : "sap-icon://error";
            oRow.exceptionState = bSuccess
              ? ValueState.Information
              : ValueState.Error;
          } else {
            oRow.productionorder = res.ProductionOrder || "";
            oRow.exceptionIcon = bSuccess
              ? "sap-icon://sys-enter-2"
              : "sap-icon://error";
            oRow.exceptionState = bSuccess
              ? ValueState.Success
              : ValueState.Error;
          }
        });

        oModel.refresh();
      },

      getCurrentFileName() {
        return this.byId("ppFileUploader").getValue() || null;
      },

      onClearUpload() {
        this._refreshTable();
        MessageToast.show(
          this._t("CLEAR_UPLOAD", null, "Đã xóa dữ liệu tải lên"),
        );
      },

      // ── History ──

      onRefreshHistory() {
        const oTable = this.byId("ppHistoryTable");
        const oBinding = oTable && oTable.getBinding("items");
        if (oBinding) oBinding.refresh();
      },

      // ── UI Utilities ──

      /**
       * Auto-size cột theo nội dung (port từ app cloud).
       */
      _autoSizeColumns(aData) {
        const oTable = this.byId(TABLE_ID);
        if (!oTable) return;

        const mColumnMap = {
          ppColProdOrder: "productionorder",
          ppColIdDoc: "id_doc",
          ppColOrderType: "ordertype",
          ppColPlant: "productionplant",
          ppColMaterial: "material",
          ppColVer: "productionversion",
          ppColQty: "totalqty",
          ppColUnit: "baseunit",
          ppColStart: "datestart",
          ppColEnd: "dateend",
          ppColSO: "saleorder",
          ppColSOItem: "saleorderitem",
          ppColNote: "longtext",
        };

        oTable.getColumns().forEach((oColumn) => {
          const sColId = oColumn.getId().split("--").pop();
          const sProp = mColumnMap[sColId];
          if (!sProp) return;

          let iMaxLen = 0;
          aData.forEach((row) => {
            const sVal = row[sProp] ? String(row[sProp]) : "";
            if (sVal.length > iMaxLen) iMaxLen = sVal.length;
          });
          const sHeader = oColumn.getLabel().getText();
          if (sHeader.length > iMaxLen) iMaxLen = sHeader.length;

          let iWidth = iMaxLen * 0.6 + 2;
          if (iWidth < 6) iWidth = 6;
          if (iWidth > 30) iWidth = 30;
          oColumn.setWidth(iWidth + "rem");
        });
      },

      _showBusy() {
        const oBusy = this.byId("idBusyDialog");
        if (oBusy) oBusy.open();
      },
      _closeBusy() {
        const oBusy = this.byId("idBusyDialog");
        if (oBusy) oBusy.close();
      },
      _destroyBusy() {
        const oBusy = this.byId("idBusyDialog");
        if (oBusy) {
          oBusy.close();
          oBusy.destroy();
        }
      },
      _refreshTable() {
        const oModel = this.getView().getModel(UPLOAD_MODEL);
        this.dataUpload = [];
        oModel.setProperty("/items", []);
        this.byId(POST_BTN_ID).setEnabled(false);
        this.byId(CHECK_BTN_ID).setEnabled(false);
        this.byId("ppFileUploader")?.clear();
      },
    });
  },
);