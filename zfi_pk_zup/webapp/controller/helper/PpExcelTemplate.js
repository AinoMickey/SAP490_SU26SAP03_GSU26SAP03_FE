sap.ui.define([], function () {
  "use strict";

  const COLUMN_DATA = {
    id_doc: "ID*\n(STT của lệnh)",
    ordertype: "Order Type*\n(Loại lệnh SX, vd: PP01)",
    productionplant: "Production Plant*\n(Nhà máy, vd: ZLA1/ZVT1)",
    material: "Material*\n(Mã thành phẩm)",
    productionversion: "Production Version*\n(Phiên bản SX, vd: 0001)",
    totalqty: "Total Quantity*\n(Số lượng)",
    baseunit: "Base Unit\n(ĐVT, trống = lấy theo material)",
    datestart: "Start Basic Date*\n(Ngày bắt đầu DD/MM/YYYY)",
    dateend: "End Basic Date*\n(Ngày kết thúc DD/MM/YYYY)",
    saleorder: "Sales Order\n(SO cho lệnh MTO, trống nếu MTS)",
    saleorderitem: "Sales Order Item\n(Item của SO)",
    longtext: "Note\n(Ghi chú)",
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
    fill: { fgColor: { rgb: "1B5E20" } },
    font: { bold: true, sz: 11, name: "Calibri", color: { rgb: "FFFFFF" } },
    alignment: { horizontal: "center", vertical: "center", wrapText: true },
    border: BORDER,
  };

  return {
    // MỚI: cho HistoryFileExport mượn lại đúng bộ key + label của template,
    // để file dựng lại từ log có header y hệt file gốc.
    getColumnData: function () {
      return COLUMN_DATA;
    },

    download: function () {
      const headerKeys = Object.keys(COLUMN_DATA);
      const headerLabels = Object.values(COLUMN_DATA);
      const sheet = XLSX.utils.aoa_to_sheet([headerKeys, headerLabels]);

      // Mở rộng range tới MAX_ROWS và set format Text cho toàn vùng
      const range = XLSX.utils.decode_range(sheet["!ref"]);
      range.e.r = MAX_ROWS - 1;
      sheet["!ref"] = XLSX.utils.encode_range(range);

      for (let R = 0; R < MAX_ROWS; ++R) {
        for (let C = range.s.c; C <= range.e.c; ++C) {
          const addr = XLSX.utils.encode_cell({ r: R, c: C });
          if (!sheet[addr]) sheet[addr] = { t: "s", v: "" };
          sheet[addr].z = "@"; // Excel format code "Text"
          sheet[addr].t = "s";
        }
      }

      // Style 2 dòng header (sau khi set format để không bị ghi đè)
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
        oNow.getFullYear() +
        sPad(oNow.getMonth() + 1) +
        sPad(oNow.getDate()) +
        "_" +
        sPad(oNow.getHours()) +
        sPad(oNow.getMinutes()) +
        sPad(oNow.getSeconds());
      XLSX.writeFile(wb, "TEMPLATE_PP_UPLSX_" + sStamp + ".xlsx");
    },
  };
});
