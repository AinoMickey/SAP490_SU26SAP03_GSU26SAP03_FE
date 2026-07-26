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
    Filter, // MỚI
    FilterOperator, // MỚI
    HistoryDetailDialog, // MỚI
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
          new JSONModel({ items: [] }),
          UPLOAD_TABLE_MODEL,
        );
        this._buildColumns();
      },

      onExit() {
        this._destroyBusy();
        ColumnSettingsDialog.destroy(this);
      },

      _getODataModel() {
        return this.getView().getModel();
      },

      // ── Dynamic Columns (88 cột từ ColumnConfig) ──

      /**
       * Sinh 88 cột từ ColumnConfig vào tableUpload.
       * Visible khởi tạo: lựa chọn đã lưu localStorage, không có thì default 13 cột.
       */
      _buildColumns() {
        const oTable = this.byId(UPLOAD_TABLE_ID);
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
        MessageToast.show("Đang hiển thị tất cả 88 field");
      },

      /** Nút Setting: mở dialog chọn cột */
      openSetting() {
        ColumnSettingsDialog.open(this, (aSelectedKeys) => {
          this._applyColumnVisibility(aSelectedKeys);
          MessageToast.show("Đã áp dụng " + aSelectedKeys.length + " cột");
        });
      },

      // ── Navigation ──

      onNavBack() {
        this.getOwnerComponent().getRouter().navTo("RouteMain");
      },

      // ── Template ──

      onDownloadTemplate() {
        ExcelTemplate.download();
        MessageToast.show("Template File Downloading...");
      },

      // ── Upload ──

      onFileChange: async function (oEvent) {
        const oFile =
          oEvent.getParameter("files") && oEvent.getParameter("files")[0];
        const oModel = this.getView().getModel(UPLOAD_MODEL); // tên model FI của anh

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

          // ══════════ MỚI 1: nạp validator + check template FI ══════════
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
                  title: "File không đúng template FI",
                  description:
                    "Thiếu cột: " +
                    oTpl.missing.join(", ") +
                    ". Vui lòng tải Template mới và điền lại.",
                },
              ],
              this,
            );
            this._refreshTable();
            return;
          }

          // ══════════ MỚI 2: bóc label an toàn (THAY dòng delete excelData[0]) ══════════
          const oStrip = UploadValidator.stripLabelRow(excelData);
          const aPrepared = UploadValidator.prepareRows(
            oStrip.rows,
            oStrip.startRow,
          );

          // thay cho check (excelData.length <= 1) cũ
          if (aPrepared.length === 0) {
            MessageToast.show("File không có dòng dữ liệu!");
            this._refreshTable();
            return;
          }

          // ══════════ MỚI 3: chặn syntax FI (row + doc-level) ══════════
          const aSynErrors = UploadValidator.validateFI(aPrepared);
          if (aSynErrors.length > 0) {
            ErrorDialog.handleErrorDialog(
              aSynErrors.map((e) => ({
                type: "Error",
                title: `Dòng Excel ${e.excelRow} - cột ${e.field}`,
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
          this._buildColumns(); // các bước sẵn có của FI (build cột động...) giữ nguyên

          MessageToast.show(
            "Upload thành công " + this.dataUpload.length + " dòng",
          );
          this.byId("postButton").setEnabled(true); // đúng id nút của FI
          this.byId("checkButton").setEnabled(true);
        } catch (error) {
          MessageBox.error(error?.message || "Upload failed or was cancelled.");
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

          const oResponse = await ApiService.callActionUpload(
            oModel,
            this.getCurrentFileName(),
            isUpdate,
            sTestMode,
            aDocs,
          );

          ResultsDialog.show(this, oResponse.messages, sTestMode, POST_BTN_ID);
        } catch (error) {
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
        MessageToast.show("Đã xóa dữ liệu upload");
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
                filters: [
                  new Filter("Filename", FilterOperator.EQ, o.filename),
                ],
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
      },
    });
  },
);
