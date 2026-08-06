sap.ui.define(
  [
    "sap/ui/core/mvc/Controller",
    "sap/ui/model/json/JSONModel",
    "sap/ui/model/Filter",
    "sap/ui/model/FilterOperator",
    "sap/m/MessageBox",
    "sap/m/MessageToast",
    "sap/m/Dialog",
    "sap/m/Button",
    "sap/m/Table",
    "sap/m/Column",
    "sap/m/ColumnListItem",
    "sap/m/Text",
    "./helper/Formatter",
    "./helper/SimpleChart",
    "./helper/ExcelExport",
    "./xlsx/xlsx.bundle",
    "sap/ui/core/Item",
  ],
  function (
    Controller,
    JSONModel,
    Filter,
    FilterOperator,
    MessageBox,
    MessageToast,
    Dialog,
    Button,
    Table,
    Column,
    ColumnListItem,
    Text,
    Formatter,
    SimpleChart,
    ExcelExport,
    xlsxBundle,
    Item,
  ) {
    "use strict";

    const PAGE_SIZE = 5000; // mỗi lượt request OData
    const HARD_LIMIT = 50000; // trần an toàn phía client, vượt thì yêu cầu thu hẹp filter

    return Controller.extend("zuprpt.controller.FiRpt", {
      formatter: Formatter,

      onInit() {
        this.getView().setModel(
          new JSONModel({
            rows: [],
            chart: [],
            kpi: {
              docCount: 0,
              fileCount: 0,
              lineCount: 0,
              updCount: 0,
              amountLC: 0,
              amountScale: "",
              amountFull: "0",
              localCur: "",
              amountSub: "",
              amountState: "None",
              curCount: 0,
              mixed: false,
              byCurText: "",
            },
          }),
          "rpt",
        );
        this.getOwnerComponent()
          .getRouter()
          .getRoute("RouteFiRpt")
          .attachPatternMatched(this._onRouteMatched, this);
      },

      onExit() {
        const oBusy = this.byId("idBusyDialog");
        if (oBusy) {
          oBusy.close();
          oBusy.destroy();
        }
        if (this._oItemDialog) {
          this._oItemDialog.destroy();
          this._oItemDialog = null;
        }
      },

      _onRouteMatched() {
        const oRange = this.byId("fiDateRange");
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
        this.byId("fiDateRange").setDateValue(null);
        this.byId("fiDateRange").setSecondDateValue(null);
        this.byId("fiCompanyCode").setValue("");
        this.byId("fiDocType").setValue("");
        this.byId("fiFilename").setValue("");
        this.byId("fiCurrency").setSelectedKey("");
      },

      _buildFilters() {
        const aFilters = [];
        const oRange = this.byId("fiDateRange");
        if (oRange.getDateValue() && oRange.getSecondDateValue()) {
          aFilters.push(
            new Filter(
              "PstDate",
              FilterOperator.BT,
              this._iso(oRange.getDateValue()),
              this._iso(oRange.getSecondDateValue()),
            ),
          );
        }
        const sCc = this.byId("fiCompanyCode").getValue().trim().toUpperCase();
        if (sCc)
          aFilters.push(new Filter("CompanyCode", FilterOperator.EQ, sCc));

        const sDt = this.byId("fiDocType").getValue().trim().toUpperCase();
        if (sDt)
          aFilters.push(new Filter("DocumentType", FilterOperator.EQ, sDt));

        const sFn = this.byId("fiFilename").getValue().trim();
        if (sFn)
          aFilters.push(new Filter("Filename", FilterOperator.Contains, sFn));
        const sCur = this.byId("fiCurrency").getSelectedKey();
        if (sCur)
          aFilters.push(new Filter("LocalCurrency", FilterOperator.EQ, sCur));
        return aFilters;
      },

      _iso(oDate) {
        const p = (n) => String(n).padStart(2, "0");
        return (
          oDate.getFullYear() +
          "-" +
          p(oDate.getMonth() + 1) +
          "-" +
          p(oDate.getDate())
        );
      },

      onGo: async function () {
        const oBusy = this.byId("idBusyDialog");
        if (oBusy) oBusy.open();
        try {
          const oModel = this.getView().getModel();
          const oBinding = oModel.bindList(
            "/FIUploadReport",
            null,
            [],
            this._buildFilters(),
            { $orderby: "PstDate desc,Filename,IdDoc", $count: true },
          );

          // 1. Lấy tổng số dòng theo filter trước khi tải
          await oBinding.requestContexts(0, 1);
          const iTotal =
            Number(
              await oBinding.getHeaderContext().requestProperty("$count"),
            ) || 0;

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
                `Vui lòng thu hẹp khoảng thời gian hoặc thêm điều kiện lọc.`,
            );
            return;
          }

          // 2. Tải tuần tự từng trang đến hết
          const aRows = [];
          for (let iStart = 0; iStart < iTotal; iStart += PAGE_SIZE) {
            if (oBusy) {
              oBusy.setText(
                `Đang tải ${Math.min(iStart + PAGE_SIZE, iTotal).toLocaleString("vi-VN")}` +
                  ` / ${iTotal.toLocaleString("vi-VN")} dòng...`,
              );
            }
            const aCtx = await oBinding.requestContexts(
              iStart,
              Math.min(PAGE_SIZE, iTotal - iStart),
            );
            aCtx.forEach((c) => aRows.push(c.getObject()));
          }

          this._applyData(aRows);
          MessageToast.show(
            `Đã tải ${aRows.length.toLocaleString("vi-VN")} chứng từ`,
          );
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
        const oFiles = new Set();
        let iLines = 0,
          iUpd = 0;

        // Tổng tiền TÁCH THEO loại tiền, không gộp
        const mAmount = {};

        aRows.forEach((r) => {
          oFiles.add(r.Filename);
          iLines += Number(r.LineCount) || 0;
          if (r.UpdDate && String(r.UpdDate).indexOf("0000") !== 0) iUpd += 1;

          const sCur = r.LocalCurrency || "?";
          mAmount[sCur] = (mAmount[sCur] || 0) + (Number(r.TotalAmountLC) || 0);
        });

        // Sắp theo giá trị tuyệt đối giảm dần, loại tiền lớn nhất đứng đầu
        const aByCur = Object.keys(mAmount)
          .map((c) => ({
            cur: c,
            amount: mAmount[c],
            amountFmt: Formatter.amount(mAmount[c]),
          }))
          .sort((a, b) => Math.abs(b.amount) - Math.abs(a.amount));

        const oTop = aByCur[0] || { cur: "", amount: 0, amountFmt: "0" };
        const bMixed = aByCur.length > 1;
        const oScaled = this._scaleAmount(oTop.amount);

        this._aRows = aRows;
        this._aByCur = aByCur; // popover và Excel dùng lại
        oRpt.setProperty("/rows", aRows);
        this._renderChart();

        oRpt.setProperty("/kpi", {
          docCount: aRows.length,
          fileCount: oFiles.size,
          lineCount: iLines,
          updCount: iUpd,

          // Chỉ là tổng của MỘT loại tiền, loại có giá trị lớn nhất
          amountLC: oScaled.value,
          amountScale: oScaled.scale,
          amountFull: oTop.amountFmt,
          localCur: oTop.cur,

          // Nói rõ đây là số của 1 loại tiền, còn bao nhiêu loại nữa
          curCount: aByCur.length,
          mixed: bMixed,
          amountSub: bMixed
            ? oTop.cur + "  ·  còn " + (aByCur.length - 1) + " loại tiền khác"
            : oTop.cur,
          amountState: bMixed ? "Critical" : "None",
          byCurText: aByCur.map((x) => x.cur + ": " + x.amountFmt).join("\n"),
        });
        const oSel = this.byId("fiCurrency");
        if (oSel && !oSel.getSelectedKey()) {
          oSel.destroyItems();
          oSel.addItem(new Item({ key: "", text: "Tất cả" }));
          aByCur.forEach((x) =>
            oSel.addItem(new Item({ key: x.cur, text: x.cur })),
          );
        }
        aRows.forEach(r => { r.TotalAmountLC = Number(r.TotalAmountLC) || 0; });
      },

      _scaleAmount(fValue) {
        const n = Number(fValue) || 0;
        const abs = Math.abs(n);
        const fmt = (x, d) =>
          x.toLocaleString("vi-VN", { maximumFractionDigits: d });
        if (abs >= 1e9) return { value: fmt(n / 1e9, 2), scale: "tỷ" };
        if (abs >= 1e6) return { value: fmt(n / 1e6, 2), scale: "tr" };
        if (abs >= 1e4) return { value: fmt(n / 1e3, 1), scale: "ng" };
        return { value: fmt(n, 2), scale: "" };
      },

      onAmountTilePress() {
        const a = this._aByCur || [];
        if (a.length <= 1) return;
        MessageBox.information(
          a.map((x) => x.cur + ":  " + x.amountFmt).join("\n"),
          { title: "Tổng tiền Nợ theo từng loại tiền" },
        );
      },

      // ===== Chart phân tích động =====

      onChartConfigChange() {
        this._renderChart();
      },

      /** Gom nhóm client-side theo dimension/measure đang chọn, top 10 + "Khác" */
      _renderChart() {
        const oHtml = this.byId("fiChartHtml");
        if (!oHtml) return;
        const aRows = this._aRows || [];
        const sDim = this.byId("fiChartDim").getSelectedKey();
        const sMea = this.byId("fiChartMeasure").getSelectedKey();
        if (
          sMea === "amount" &&
          (this._aByCur || []).length > 1 &&
          sDim !== "currency"
        ) {
          MessageToast.show(
            "Dữ liệu đang có nhiều loại tiền. Chọn Nhóm theo = Loại tiền, " +
              "hoặc lọc về một loại tiền, thì số tiền mới có nghĩa.",
          );
        }

        const fnKey = {
          doctype: (r) => r.DocumentType || "(trống)",
          company: (r) => r.CompanyCode || "(trống)",
          month: (r) =>
            r.PstDate ? String(r.PstDate).substring(0, 7) : "(trống)",
          currency: (r) => r.LocalCurrency || "(trống)",
        }[sDim];
        const fnVal = {
          count: () => 1,
          amount: (r) => Number(r.TotalAmountLC) || 0,
        }[sMea];

        const m = {};
        aRows.forEach((r) => {
          const k = fnKey(r);
          m[k] = (m[k] || 0) + fnVal(r);
        });

        let aData = Object.keys(m).map((k) => ({
          label:
            sDim === "month" && k.length === 7
              ? k.substring(5, 7) + "/" + k.substring(0, 4)
              : k,
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

      // ===== Drill-down bút toán =====

      onShowItems: async function (oEvent) {
        const oCtx = oEvent.getParameter("row").getBindingContext("rpt");
        if (!oCtx) return;
        const oDoc = oCtx.getObject();
        const oBusy = this.byId("idBusyDialog");
        if (oBusy) oBusy.open();
        try {
          const oModel = this.getView().getModel();
          const oBinding = oModel.bindList(
            "/FIUploadItem",
            null,
            [],
            [
              new Filter("Filename", FilterOperator.EQ, oDoc.Filename),
              new Filter("IdDoc", FilterOperator.EQ, oDoc.IdDoc),
            ],
            { $orderby: "IdLine" },
          );
          const aCtx = await oBinding.requestContexts(0, 999);
          const aItems = aCtx.map((c) => c.getObject());
          this._openItemDialog(oDoc, aItems);
        } catch (e) {
          MessageBox.error("Lỗi tải bút toán: " + (e.message || e));
        } finally {
          if (oBusy) oBusy.close();
        }
      },

      _openItemDialog(oDoc, aItems) {
        if (!this._oItemDialog) {
          const oTable = new Table({
            columns: [
              new Column({ header: new Text({ text: "Line" }) }),
              new Column({ header: new Text({ text: "PK" }) }),
              new Column({ header: new Text({ text: "Account" }) }),
              new Column({ header: new Text({ text: "KH/NCC" }) }),
              new Column({
                header: new Text({ text: "Amount LC" }),
                hAlign: "End",
              }),
              new Column({ header: new Text({ text: "Cur" }) }),
              new Column({ header: new Text({ text: "Cost Center" }) }),
              new Column({ header: new Text({ text: "Item Text" }) }),
            ],
            items: {
              path: "itm>/items",
              template: new ColumnListItem({
                cells: [
                  new Text({ text: "{itm>IdLine}" }),
                  new Text({ text: "{itm>PostingKey}" }),
                  new Text({ text: "{itm>Account}" }),
                  new Text({ text: "{= ${itm>Customer} || ${itm>Supplier} }" }),
                  new Text({
                    text: { path: "itm>AmountLC", formatter: Formatter.amount },
                  }),
                  new Text({ text: "{itm>LocalCurrency}" }),
                  new Text({ text: "{itm>CostCenter}" }),
                  new Text({ text: "{itm>ItemText}" }),
                ],
              }),
            },
          });
          this._oItemDialog = new Dialog({
            contentWidth: "60rem",
            contentHeight: "24rem",
            resizable: true,
            draggable: true,
            content: [oTable],
            endButton: new Button({
              text: "Đóng",
              press: () => this._oItemDialog.close(),
            }),
          });
          this._oItemDialog.setModel(new JSONModel({ items: [] }), "itm");
          this.getView().addDependent(this._oItemDialog);
        }
        this._oItemDialog.setTitle(
          `Bút toán - File ${oDoc.Filename} / ID ${oDoc.IdDoc}` +
            (oDoc.AccountingDocument ? ` / CT ${oDoc.AccountingDocument}` : ""),
        );
        this._oItemDialog.getModel("itm").setProperty("/items", aItems);
        this._oItemDialog.open();
      },

      // ===== Export =====

      onExport() {
        const oRpt = this.getView().getModel("rpt");
        const aRows = oRpt.getProperty("/rows");
        if (!aRows.length) {
          MessageToast.show("Không có dữ liệu để xuất");
          return;
        }
        const oKpi = oRpt.getProperty("/kpi");
        const oRange = this.byId("fiDateRange");

        ExcelExport.export({
          title: "BÁO CÁO CHỨNG TỪ FI ĐÃ UPLOAD",
          filePrefix: "FI_UPLOAD_REPORT",
          filters: [
            {
              label: "Posting Date (log)",
              value: oRange.getDateValue()
                ? Formatter.date(this._iso(oRange.getDateValue())) +
                  " - " +
                  Formatter.date(this._iso(oRange.getSecondDateValue()))
                : "",
            },
            {
              label: "Company Code",
              value: this.byId("fiCompanyCode").getValue(),
            },
            {
              label: "Document Type",
              value: this.byId("fiDocType").getValue(),
            },
            {
              label: "Filename chứa",
              value: this.byId("fiFilename").getValue(),
            },
          ],
          kpis: [
            { label: "Tổng chứng từ", value: oKpi.docCount },
            { label: "Tổng file", value: oKpi.fileCount },
            { label: "Tổng dòng bút toán", value: oKpi.lineCount },
            { label: "Đã update lại", value: oKpi.updCount },
            ...(this._aByCur || []).map((x) => ({
              label: "Tổng tiền Nợ (" + x.cur + ")",
              value: x.amountFmt,
            })),
          ],
          columns: [
            { key: "IdDoc", label: "ID", width: 8 },
            { key: "AccountingDocument", label: "Accounting Doc", width: 14 },
            { key: "FiscalYear", label: "Fiscal Year", width: 10 },
            { key: "CompanyCode", label: "Company", width: 10 },
            { key: "DocumentType", label: "Doc Type", width: 9 },
            { key: "DocumentDate", label: "Doc Date", width: 12, type: "date" },
            {
              key: "PostingDate",
              label: "Posting Date",
              width: 12,
              type: "date",
            },
            {
              key: "TotalAmountLC",
              label: "Tiền Nợ (LC)",
              width: 15,
              type: "qty",
            },
            { key: "LocalCurrency", label: "LC Cur", width: 7 },
            { key: "Currency", label: "Doc Cur", width: 8 },
            { key: "HeaderText", label: "Header Text", width: 28 },
            { key: "LineCount", label: "Lines", width: 7, type: "qty" },
            { key: "Filename", label: "Filename", width: 34 },
            { key: "PstDate", label: "Pst Date", width: 12, type: "date" },
            { key: "PstUser", label: "Posted By", width: 12 },
            { key: "UpdDate", label: "Upd Date", width: 12, type: "date" },
            { key: "UpdUser", label: "Upd By", width: 12 },
          ],
          rows: aRows,
        });
        MessageToast.show("Đang xuất Excel...");
      },
    });
  },
);
