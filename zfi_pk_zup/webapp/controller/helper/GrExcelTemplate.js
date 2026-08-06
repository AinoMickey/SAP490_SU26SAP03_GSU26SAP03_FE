sap.ui.define([], function () {
  "use strict";

  const COLUMN_DATA = {
    gr_number: "GR Number*\n(Tự đặt, nhiều dòng cùng GR Number sẽ gộp thành 1 phiếu)",
    document_date: "Document Date*\n(Ngày chứng từ DD/MM/YYYY)",
    movement_type: "Movement Type\n(Trống = mặc định 101)",
    po_number: "PO Number*\n(Purchase Order đã release)",
    po_item: "PO Item*\n(Item của PO)",
    receive_qty: "Receive Qty*\n(Số lượng thực nhận, <= Open Qty)",
    unit: "Unit*\n(Phải khớp Order Unit của PO item)",
    batch: "Batch",
    storage_location: "Storage Location*\n(Sloc nhận hàng)",
  };

  const MAX_ROWS = 1000;

  const BORDER = {
    top: { style: "thin", color: { rgb: "000000" } },
    bottom: { style: "thin", color: { rgb: "000000" } },
    left: { style: "thin", color: { rgb: "000000" } },
    right: { style: "thin", color: { rgb: "000000" } },
  };
  const KEY_STYLE = {
    fill: { fgColor: { rgb: "EAEAEA" } },
    font: { bold: true, sz: 11, name: "Calibri" },
    alignment: { horizontal: "center", vertical: "center" },
    border: BORDER,
  };
  const LABEL_STYLE = {
    fill: { fgColor: { rgb: "0553CE" } },
    font: { bold: true, sz: 11, name: "Calibri", color: { rgb: "FFFFFF" } },
    alignment: { horizontal: "center", vertical: "center", wrapText: true },
    border: BORDER,
  };

  return {
    download: function () {
      const headerKeys = Object.keys(COLUMN_DATA);
      const headerLabels = Object.values(COLUMN_DATA);
      const mappingHint = [
        [
          "gr_number", "document_date", "movement_type", "po_number",
          "po_item", "receive_qty", "unit", "batch", "storage_location"
        ],
        [
          "GR Number", "Document Date", "Movement Type", "PO Number",
          "PO Item", "Receive Qty", "Unit",  "batch", "Storage Location"
        ],
        [
          "Source column must match these fields", "Định dạng date DD/MM/YYYY",
          "Mặc định 101 nếu để trống", "PO đã release",
          "Item của PO", "Số lượng thực nhận", "Đơn vị theo PO", "Mã Lô" , "Storage location"
        ]
      ];
      const sheet = XLSX.utils.aoa_to_sheet([headerKeys, headerLabels]);

      const range = XLSX.utils.decode_range(sheet["!ref"]);
      range.e.r = MAX_ROWS - 1;
      sheet["!ref"] = XLSX.utils.encode_range(range);

      for (let R = 0; R < MAX_ROWS; ++R) {
        for (let C = range.s.c; C <= range.e.c; ++C) {
          const addr = XLSX.utils.encode_cell({ r: R, c: C });
          if (!sheet[addr]) sheet[addr] = { t: "s", v: "" };
          sheet[addr].z = "@";
          sheet[addr].t = "s";
        }
      }

      const cols = [];
      for (let C = range.s.c; C <= range.e.c; ++C) {
        cols.push({ wch: 22 });
        sheet[XLSX.utils.encode_cell({ c: C, r: 0 })].s = KEY_STYLE;
        sheet[XLSX.utils.encode_cell({ c: C, r: 1 })].s = LABEL_STYLE;
      }
      sheet["!cols"] = cols;
      sheet["!rows"] = [{ hpx: 20 }, { hpx: 50 }];

      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, sheet, "Data");
      const oNow = new Date();
      const sPad = (n) => String(n).padStart(2, "0");
      const sStamp =
        oNow.getFullYear() + sPad(oNow.getMonth() + 1) + sPad(oNow.getDate()) +
        "_" + sPad(oNow.getHours()) + sPad(oNow.getMinutes()) + sPad(oNow.getSeconds());
        const hintSheet = XLSX.utils.aoa_to_sheet(mappingHint);
      hintSheet["!cols"] = headerKeys.map(() => ({ wch: 28 }));
      XLSX.utils.book_append_sheet(wb, hintSheet, "Mapping");
      XLSX.writeFile(wb, "TEMPLATE_MM_POGR_" + sStamp + ".xlsx");

    },

        downloadPrefilled: function (aRows) {
      const headerKeys = Object.keys(COLUMN_DATA);
      const headerLabels = Object.values(COLUMN_DATA);
      const aData = aRows.map((row) => ({
        gr_number: "",
        document_date: "",
        movement_type: "101",
        po_number: row.PurchaseOrder || "",
        po_item: row.PurchaseOrderItem || "",
        receive_qty: "",
        unit: row.OrderUnit || "",
        batch: "",
        storage_location: row.StorageLocation || "",
      }));
      const aoa = [headerKeys, headerLabels, ...aData.map((r) => headerKeys.map((k) => r[k]))];
      const sheet = XLSX.utils.aoa_to_sheet(aoa);

      const range = XLSX.utils.decode_range(sheet["!ref"]);
      for (let R = 0; R <= range.e.r; ++R) {
        for (let C = range.s.c; C <= range.e.c; ++C) {
          const addr = XLSX.utils.encode_cell({ r: R, c: C });
          if (!sheet[addr]) sheet[addr] = { t: "s", v: "" };
          sheet[addr].z = "@";
          sheet[addr].t = "s";
        }
      }
      const cols = [];
      for (let C = range.s.c; C <= range.e.c; ++C) {
        cols.push({ wch: 22 });
        sheet[XLSX.utils.encode_cell({ c: C, r: 0 })].s = KEY_STYLE;
        sheet[XLSX.utils.encode_cell({ c: C, r: 1 })].s = LABEL_STYLE;
      }
      sheet["!cols"] = cols;

      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, sheet, "Data");
      const oNow = new Date();
      const sPad = (n) => String(n).padStart(2, "0");
      const sStamp =
        oNow.getFullYear() + sPad(oNow.getMonth() + 1) + sPad(oNow.getDate()) +
        "_" + sPad(oNow.getHours()) + sPad(oNow.getMinutes()) + sPad(oNow.getSeconds());
      XLSX.writeFile(wb, "GR_FROM_PO_" + sStamp + ".xlsx");
    },

    getColumnData: function () {
      return Object.assign({}, COLUMN_DATA);
    },

  };
});
