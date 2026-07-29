sap.ui.define(
  [
    "sap/ui/core/mvc/Controller",
    "sap/m/MessageToast",
    "sap/ui/model/json/JSONModel",
    "sap/m/MessageBox",
    "sap/m/Label",
    "sap/m/Text",
    "sap/ui/table/Column",
    "./helper/ExcelTemplate",
    "./helper/ExcelParser",
    "./helper/ApiService",
    "./helper/ResultsDialog",
    "./helper/ErrorDialog",
    "./helper/ColumnConfig",
    "./helper/ColumnSettingsDialog",
    "./xlsx/xlsx.bundle",
    "sap/ui/model/Filter",
    "sap/ui/model/FilterOperator",
    "./helper/HistoryDetailDialog",
  ],
  function (
    Controller,
    MessageToast,
    JSONModel,
    MessageBox,
    Label,
    Text,
    Column,
    ExcelTemplate,
    ExcelParser,
    ApiService,
    ResultsDialog,
    ErrorDialog,
    ColumnConfig,
    ColumnSettingsDialog,
    xlsxBundle, // SỬA: trước đây THIẾU tham số này -> 18 dependency / 17 tham số,
    Filter, //     làm Filter/FilterOperator/HistoryDetailDialog trỏ lệch 1 ô.
    FilterOperator, //  Chưa nổ vì mọi chỗ dùng đều nằm trong sap.ui.require lồng bên trong.
    HistoryDetailDialog,
  ) {
    "use strict";

    const UPLOAD_TABLE_MODEL = "uploadModel";
    const UPLOAD_TABLE_ID = "tableUpload";
    const POST_BTN_ID = "postButton";
    const CHECK_BTN_ID = "checkButton";

    return Controller.extend("zfipkzup.controller.Fi", {
      dataUpload: [],

      onInit() {
        this.getView().setModel(
          new JSONModel({
            items: [],
            // MỚI: trạng thái post hiện trên toolbar, còn lại sau khi đóng dialog
            postState: "None", // None | Success | Warning | Error | Information
            postText: this._t("STATUS_NOT_POSTED", null, "Chưa post"),
            postDocs: "",
          }),
          UPLOAD_TABLE_MODEL,
        );
        this._buildColumns();
        //ThaoNTT add
        this._loadAuth();

      },

      onExit() {
        this._destroyBusy();
        ColumnSettingsDialog.destroy(this);
      },

      _getODataModel() {
        return this.getView().getModel();
      },

      _getBundle() {
        const oModel = this.getView().getModel("i18n");
        return oModel ? oModel.getResourceBundle() : null;
      },

      /**
       * getText trả về CHÍNH KEY khi thiếu entry -> `getText(k) || fallback`
       * không bao giờ chạy tới fallback mà đẩy tên key ra màn hình.
       * Hàm này so với key nên fallback mới có tác dụng.
       */
      _t(sKey, aArgs, sFallback) {
        const oBundle = this._getBundle();
        if (!oBundle) return sFallback;
        const s = oBundle.getText(sKey, aArgs);
        return !s || s === sKey ? sFallback : s;
      },

      // ── Dynamic Columns (88 cột từ ColumnConfig) ──

      /**
       * Sinh 88 cột từ ColumnConfig vào tableUpload.
       * Visible khởi tạo: lựa chọn đã lưu localStorage, không có thì default 13 cột.
       */
      _buildColumns() {
        const oTable = this.byId(UPLOAD_TABLE_ID);
        oTable.destroyColumns();
        const aVisibleKeys = ColumnSettingsDialog.getInitialVisibleKeys();

        ColumnConfig.getColumns().forEach((c) => {
          const oColumn = new Column(this.createId("col_" + c.key), {
            width: c.width || "10rem",
            visible: aVisibleKeys.indexOf(c.key) !== -1,
            label: new Label({ text: c.label, wrapping: false }),
            template: new Text({
              text: "{" + UPLOAD_TABLE_MODEL + ">" + c.key + "}",
              wrapping: false,
            }),
          });
          // Gắn key để dialog đọc lại trạng thái hiện tại của bảng
          oColumn.data("colKey", c.key);
          oTable.addColumn(oColumn);
        });
      },

      /** Áp visibility theo danh sách key được chọn */
      _applyColumnVisibility(aVisibleKeys) {
        const oTable = this.byId(UPLOAD_TABLE_ID);
        oTable.getColumns().forEach((oCol) => {
          const sKey = oCol.data("colKey");
          oCol.setVisible(aVisibleKeys.indexOf(sKey) !== -1);
        });
      },

      /** Nút All Field: hiện đủ 88 cột */
      onShowAllColumns() {
        const aAll = ColumnConfig.getAllKeys();
        this._applyColumnVisibility(aAll);
        ColumnSettingsDialog.saveVisibleKeys(aAll);
        MessageToast.show(
          this._t("P13nHelper_ShowAll", [aAll.length],
            "Đang hiển thị tất cả " + aAll.length + " field"),
        );
      },

      /** Nút Setting: mở dialog chọn cột */
      openSetting() {
        ColumnSettingsDialog.open(this, (aSelectedKeys) => {
          this._applyColumnVisibility(aSelectedKeys);
          MessageToast.show(
            this._t("APPLY_COLUMNS", [aSelectedKeys.length],
              "Đã áp dụng " + aSelectedKeys.length + " cột"),
          );
        });
      },

      // ── Navigation ──

      onNavBack() {
        this.getOwnerComponent().getRouter().navTo("RouteMain");
      },

      // ── Template ──

      onDownloadTemplate() {
        ExcelTemplate.download();
        MessageToast.show(
          this._t("TEMPLATE_LOADING", null, "Đang tải template..."),
        );
      },

      // ── Upload ──

      onFileChange: async function (oEvent) {
        const oFile =
          oEvent.getParameter("files") && oEvent.getParameter("files")[0];
        const oModel = this.getView().getModel(UPLOAD_TABLE_MODEL);

        if (!oFile) {
          this._refreshTable();
          return;
        }

        this.dataUpload = [];
        oModel.setProperty("/items", []);
        this._showBusy();

        try {
          // check trùng filename (FI có sẵn từ đầu) - GIỮ NGUYÊN
          await ApiService.checkFileExists(
            this.getView().getModel(),
            oFile.name,
          );

          const fileContent = await ExcelParser.readFile(oFile);
          const workbook = XLSX.read(fileContent, { type: "binary" });
          const oSheet =
            workbook.Sheets["Data"] || workbook.Sheets[workbook.SheetNames[0]];
          const excelData = XLSX.utils.sheet_to_row_object_array(oSheet);

          // ══════════ 1: nạp validator + check template FI ══════════
          const UploadValidator = await new Promise((res) =>
            sap.ui.require(["zfipkzup/controller/helper/UploadValidator"], res),
          );

          const aHeader =
            XLSX.utils.sheet_to_json(oSheet, { header: 1 })[0] || [];
          const oTpl = UploadValidator.checkTemplate(aHeader, "FI");
          if (!oTpl.ok) {
            ErrorDialog.handleErrorDialog(
              [
                {
                  type: "Error",
                  title: this._t("EXCEL_TEMPLATE_INVALID", ["FI"],
                    "File không đúng template FI"),
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

          // ══════════ 2: bóc label an toàn (THAY dòng delete excelData[0]) ══════════
          const oStrip = UploadValidator.stripLabelRow(excelData);
          const aPrepared = UploadValidator.prepareRows(
            oStrip.rows,
            oStrip.startRow,
          );

          if (aPrepared.length === 0) {
            MessageToast.show(
              this._t("UPLOAD_NO_DATA", null, "File không có dòng dữ liệu!"),
            );
            this._refreshTable();
            return;
          }

          // ══════════ 3: chặn syntax FI (row + doc-level) ══════════
          const aSynErrors = UploadValidator.validateFI(aPrepared);
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

          // ══════════ FLOW CŨ GIỮ NGUYÊN, chỉ đổi nguồn data ══════════
          const result = ExcelParser.processExcelData(
            [null].concat(oStrip.rows),
          );
          this.dataUpload = result.data;
          oModel.setProperty("/items", this.dataUpload);
          this._buildColumns();

          // MỚI: nạp file mới -> xóa trạng thái post của lần trước
          this._resetPostStatus();

          MessageToast.show(
            this._t("UPLOAD_SUCCESS", [this.dataUpload.length],
              "Tải lên thành công " + this.dataUpload.length + " dòng"),
          );
          this.byId(POST_BTN_ID).setEnabled(true);
          this.byId(CHECK_BTN_ID).setEnabled(true);
        } catch (error) {
          MessageBox.error(
            error?.message ||
            this._t("UPLOAD_CANCELLED", null,
              "Tải lên bị hủy hoặc thất bại."),
          );
        } finally {
          this._closeBusy();
        }
      },

      onCheck(oEvent) {
        ApiService.checkFileExists(
          this._getODataModel(),
          this.getCurrentFileName(),
        )
          .then(() => this.callApiFiDoc(oEvent, "X"))
          .catch((error) => MessageBox.error(error.message));
      },

      onPost: async function (oEvent) {
        try {
          await ApiService.checkFileExists(
            this._getODataModel(),
            this.getCurrentFileName(),
          );
          this.callApiFiDoc(oEvent, "");
        } catch (error) {
          MessageBox.error(error.message);
        }
      },

      callApiFiDoc: async function (oEvent, sTestMode) {
        this._showBusy();
        try {
          const oModel = this._getODataModel();
          const groupedDocs = ApiService.groupDataByDocId(this.dataUpload);
          const aDocs = ApiService.buildAllDocs(groupedDocs);
          const isUpdate = this.isDocumentPosted ? "X" : "";
          // Chặn double-click: tắt nút trong lúc gọi
          this.byId(POST_BTN_ID)?.setEnabled(false);
          this.byId(CHECK_BTN_ID)?.setEnabled(false);

          const oResponse = await ApiService.callActionUpload(
            oModel,
            this.getCurrentFileName(),
            isUpdate,
            sTestMode,
            aDocs,
          );

          // MỚI: truyền thêm results để dialog liệt kê được số chứng từ,
          //      nhận về bản tóm tắt để đổ ra ObjectStatus trên toolbar
          const oSum = ResultsDialog.show(
            this,
            oResponse.messages,
            sTestMode,
            POST_BTN_ID,
            oResponse.results,
          );
          this._applyPostStatus(oSum, sTestMode);
        } catch (error) {
          // Lỗi gọi service thì mở lại nút để thử lại,
          // không bắt người dùng chọn lại file
          const bHasData = this.dataUpload.length > 0;
          this.byId(CHECK_BTN_ID)?.setEnabled(bHasData);
          this.byId(POST_BTN_ID)?.setEnabled(bHasData);
          MessageBox.error(error.message || JSON.stringify(error));
        } finally {
          this._closeBusy();
        }
      },

      getCurrentFileName() {
        return this.byId("fileUploader").getValue() || null;
      },

      /** Nút Clear: xóa sạch dữ liệu upload, reset trạng thái nút */
      onClearUpload() {
        this._refreshTable();
        this.isDocumentPosted = false;
        MessageToast.show(
          this._t("CLEAR_UPLOAD", null, "Đã xóa dữ liệu tải lên"),
        );
      },

      // ── Trạng thái post (MỚI) ──

      /** Đổ kết quả Check/Post ra ObjectStatus trên toolbar */
      _applyPostStatus(oSum, sTestMode) {
        const oModel = this.getView().getModel(UPLOAD_TABLE_MODEL);
        if (!oModel || !oSum) return;

        let sText;
        if (sTestMode === "X") {
          sText =
            oSum.err > 0
              ? this._t("STATUS_CHECK_ERR", [oSum.err, oSum.total],
                "Kiểm tra: " + oSum.err + "/" + oSum.total + " lỗi")
              : this._t("STATUS_CHECK_OK", [oSum.ok, oSum.total],
                "Kiểm tra: " + oSum.ok + "/" + oSum.total + " hợp lệ");
        } else if (oSum.err === 0 && oSum.ok > 0) {
          sText = this._t("STATUS_POSTED", [oSum.ok, oSum.total],
            "Đã post " + oSum.ok + "/" + oSum.total + " chứng từ");
        } else if (oSum.ok > 0) {
          sText = this._t("STATUS_PARTIAL", [oSum.ok, oSum.total],
            "Post một phần " + oSum.ok + "/" + oSum.total);
        } else {
          sText = this._t("STATUS_FAILED", [oSum.total],
            "Post thất bại 0/" + oSum.total);
        }

        oModel.setProperty("/postState", oSum.state || "None");
        oModel.setProperty("/postText", sText);
        oModel.setProperty("/postDocs", oSum.docList || "");
      },

      /** Về trạng thái chưa post */
      _resetPostStatus() {
        const oModel = this.getView().getModel(UPLOAD_TABLE_MODEL);
        if (!oModel) return;
        oModel.setProperty("/postState", "None");
        oModel.setProperty(
          "/postText",
          this._t("STATUS_NOT_POSTED", null, "Chưa post"),
        );
        oModel.setProperty("/postDocs", "");
      },

      // ── History ──

      onRefreshHistory() {
        const oTable = this.byId("historyTable");
        const oBinding = oTable && oTable.getBinding("items");
        if (oBinding) oBinding.refresh();
      },

      onHistoryItemPress: function (oEvent) {
        const oCtx = oEvent.getSource().getBindingContext();
        if (!oCtx) return;
        const o = oCtx.getObject();
        const oRptModel = this.getOwnerComponent().getModel("rpt");
        const oView = this.getView();

        sap.ui.require(
          [
            "zfipkzup/controller/helper/HistoryDetailDialog",
            "sap/ui/model/Filter",
            "sap/ui/model/FilterOperator",
          ],
          async function (HistoryDetailDialog, Filter, FilterOperator) {
            // Context bảng chỉ có field đang bind (autoExpandSelect),
            // nên lấy chi tiết đầy đủ từ service RPT theo Filename
            let aDocs = [];
            if (oRptModel) {
              try {
                const oBinding = oRptModel.bindList(
                  "/FIUploadReport",
                  null,
                  [],
                  [new Filter("Filename", FilterOperator.EQ, o.filename)],
                  { $orderby: "IdDoc" },
                );
                const aCtx = await oBinding.requestContexts(0, 200);
                aDocs = aCtx.map((c) => c.getObject());
              } catch (e) {
                /* rpt lỗi thì dialog vẫn mở phần file */
              }
            }

            const sDocList = aDocs
              .map((d) => d.AccountingDocument)
              .filter(Boolean)
              .join(", ");

            HistoryDetailDialog.open({
              view: oView,
              title: "Chi tiết file upload",
              record: {
                filename: o.filename,
                pst_date: o.pst_date,
                pst_user: o.pst_user,
                doccount: aDocs.length || "",
                doclist: sDocList,
                companycode: aDocs[0] && aDocs[0].CompanyCode,
              },
              fields: [
                { key: "filename", label: "Filename" },
                { key: "pst_date", label: "Ngày post", type: "date" },
                { key: "pst_user", label: "Người post" },
                { key: "companycode", label: "Company Code" },
                { key: "doccount", label: "Số chứng từ trong file" },
                { key: "doclist", label: "Danh sách số CT" },
              ],
              copy: { label: "Copy số CT", value: sDocList },
              items: oRptModel && {
                model: oRptModel,
                path: "/FIUploadItem",
                filters: [new Filter("Filename", FilterOperator.EQ, o.filename)],
                sorterParams: { $orderby: "IdDoc,IdLine" },
                title: "Bút toán",
                columns: [
                  { key: "IdDoc", label: "ID" },
                  { key: "IdLine", label: "Line" },
                  { key: "PostingKey", label: "PK" },
                  { key: "Account", label: "Account" },
                  { key: "AmountLC", label: "Amount LC", type: "num" },
                  { key: "LocalCurrency", label: "Cur" },
                  { key: "ItemText", label: "Item Text" },
                ],
              },
            });
          },
        );
      },

      onDownloadHistoryFile: function (oEvent) {
        const oCtx = oEvent.getSource().getBindingContext();
        if (!oCtx) return;

        const sFilename = oCtx.getObject().filename;
        const oFiModel = this.getView().getModel(); // model mặc định = service FI
        const that = this;

        this._showBusy();

        sap.ui.require(
          ["zfipkzup/controller/helper/HistoryFileExport"],
          async function (HistoryFileExport) {
            try {
              const oRes = await HistoryFileExport.exportFi(
                oFiModel,
                sFilename,
                { withRefs: true },
              );
              MessageToast.show(
                that._t("CREATE_HISTORY_FILE", [oRes.file, oRes.rows],
                  "Đã tạo " + oRes.file + " (" + oRes.rows + " dòng)"),
              );
            } catch (e) {
              // SỬA: trong callback function thường, `this` KHÔNG phải controller
              // -> bản cũ gọi this.getView() ở đây ném TypeError và nuốt mất lỗi gốc.
              MessageBox.error(
                that._t("ERROR_BUILD_FILE", [e.message || e],
                  "Không dựng được file: " + (e.message || e)) +
                "\n\n" +
                that._t("ERROR_BUILD_FILE_HINT", null,
                  "Kiểm tra: CDS ZFI_I_DIS_UP_I đã thêm field mới và đã publish lại chưa."),
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

      // ── UI Utilities ──

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
        const oModel = this.getView().getModel(UPLOAD_TABLE_MODEL);
        this.dataUpload = [];
        oModel.setProperty("/items", []);
        this.byId(POST_BTN_ID).setEnabled(false);
        this.byId(CHECK_BTN_ID).setEnabled(false);
        this.byId("fileUploader")?.clear();
        this._resetPostStatus(); // MỚI
      },
      //ThaoNTT add
      _loadAuth: async function () {
        try {
          const oAuth = await ApiService.loadMyAuth(
            this.getOwnerComponent().getModel("gr")
          );
          if (oAuth.can_upload_fi === false) {
            this.byId(POST_BTN_ID)?.setVisible(false);
            this.byId(CHECK_BTN_ID)?.setVisible(false);
            this.byId("fileUploader")?.setEnabled(false);   // ← THÊM: khóa nút chọn file
            this._bNoAuth = true;
          }
        } catch (e) {
          // Không đọc được quyền thì để nguyên nút — backend vẫn chặn
        }
      },

    });
  },
);