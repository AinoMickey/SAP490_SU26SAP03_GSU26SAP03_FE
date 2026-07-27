sap.ui.define([], function () {
  "use strict";

  const COLUMN_DATA = {
    id_doc: "ID*\n(STT của c.từ)",
    documentdate: "Document Date*\n(Ngày hoá đơn, chứng từ)",
    postingdate: "Posting Date*\n(Ngày hạch toán)",
    documenttype: `Document Type*\n(Loại c.từ)\n+ DR: Tăng công nợ khách hàng\n+ DG: Giảm công nợ khách hàng\n+ KR: Tăng công nợ nhà cung cấp\n+ KG: Giảm công nợ nhà cung cấp\n+ SA: Chứng từ kế toán\n+ A1: Thu tiền mặt\n+ A2: Chi tiền mặt\n+ A3: Thu tiền gửi KKH\n+ A4: Chi tiền gửi KKH`,
    companycode: "Company Code\n(Mã cty I001/P001)",
    currency: "Currency Key*\n(Loại tiền)",
    exchangerate: "Exchange rate\n(Tỷ giá)",
    headertext: "Header Text\n(Diễn giải chung của c.từ, 25 ký tự)",
    referencedoc: "Reference Doc\n(Ký hiệu HĐ.Số HĐ)",
    headerref1: "Header Reference key 1",
    negativeposting: "Is Negative Posting",
    postingkey: `Posting Key*\n(Mã Nợ/Có)\n+ 40: Nợ (tài khoản không đối tượng)\n+ 50: Có (tài khoản không đối tượng)\n+ 01: Nợ (tài khoản có đối tượng - Khách hàng)\n+ 11: Có (tài khoản có đối tượng - Khách hàng)\n+ 21: Nợ (tài khoản có đối tượng - Nhà cung cấp)\n+ 31: Có (tài khoản có đối tượng - Nhà cung cấp)\n+ 70: Nợ (tài khoản có đối tượng - Tài sản)\n+ 75: Có (tài khoản có đối tượng - Tài sản)`,
    account: "Account*\n(Mã NCC/NV/KH/\nTài khoản không đối tượng)",
    mainassetnumber: "Main Asset Number",
    subassetnumber: "Sub Asset Number",
    specialglaccount: "Special GL Indicator",
    assettransactiontype: "Asset Transaction Type",
    amountindoumentcurrency:
      "Amount in Document Currency*\n(Số tiền theo c.từ)",
    amountinlocalcurrency:
      "Amount in Local Currency*\n(Số tiền đơn vị tiền tệ địa phương)",
    taxbaseamount:
      "Tax Base amount in\nDocument Currency\n(Số tiền trước VAT theo c.từ)",
    localtaxbaseamount:
      "Tax Base amount in\nLocal Currency\n(Số tiền trước VAT theo đơn vị tiền tệ địa phương)",
    assignment: "Assignment",
    costcenter: "Cost center\n(Phòng ban)",
    profitcenter: "Profit Center",
    internalorder: "Internal Order\n(Lệnh nội bộ - CO assignment)",
    wbselement: "WBS Element",
    businessarea: "Business Area",
    assetvaluedate: "Asset value date\n(Ngày tăng nguyên giá Tài sản)",
    itemtext: "Item Text\n(Diễn giải ngắn của c.từ, 50 ký tự)",
    longtext: "Long Text\n(Diễn giải dài, tối đa 100 ký tự)",
    overrideglaccount: "Override GL Account\n(TK thay thế)",
    taxcode: "Tax Code\n(Loại thuế suất)",
    segment: "Segment",
    paymentterms: "Payment term",
    paymentblockreason: "Payment Block",
    paymentmethod: "Payment Method",
    baselinedate: "Baseline date\n(Ngày bắt đầu tính công nợ)",
    valuedate: "Value Date",
    netduedate: "Net Due Date\n(Ngày đến hạn thanh toán)",
    contractnumber: "Contract Number",
    contracttype: "Contract Type",
    housebank: "House bank",
    bankaccountid: "Bank Account ID",
    invoicerefnum: "Invoice Ref. Number",
    invoicereffiscalyear: "Invoice Ref. Fiscal Year",
    invoicereflineitem: "Invoice Ref. Line Item",
    purchasingno: "Purchasing Doc. No.",
    purchasingitem: "Purchasing Doc. Item",
    saleorder: "Sales Order Number",
    saleorderitem: "Sales Order Item",
    customer: "Customer",
    cusgroup: "Customer Group",
    division: "Division",
    distributionchannel: "Distribution Channel",
    salesorganization: "Sales Organization",
    salesoffice: "Sales Office",
    salesemployee: "Sales Employee",
    salesgroup: "Sales Group",
    materialgroup: "Material Group",
    product: "Product",
    unit: "Unit of Measure\n(Đơn vị tính)",
    baseunit: "Base Unit",
    quantity: "Quantity\n(Số lượng)",
    alternativepayee: "Alternative Payee\n(Tài khoản hạch toán thay thế)",
    name1: "Name 1\n(Tên 1 của mã vãng lai)",
    name2: "Name 2\n(Tên 2 của mã vãng lai)",
    name3: "Name 3\n(Tên 3 của mã vãng lai)",
    name4: "Name 4\n(Tên 4 của mã vãng lai)",
    mst: "VAT Reg. No.\n(MST của mã vãng lai)",
    city: "City\n(Tỉnh của mã vãng lai)",
    country: "Country\n(Quốc gia của mã vãng lai)",
    vatregno: "VAT Reg. No.",
    ref1: "Ref1",
    ref2: "Ref2",
    ref3: "Ref3",
    namecus1: "Name 1\n(Tên 1 của đối tượng nhận tiền)",
    namecus2: "Name 2\n(Tên 2 của đối tượng nhận tiền)",
    namecus3: "Name 3\n(Tên 3 của đối tượng nhận tiền)",
    namecus4: "Name 4\n(Tên 4 của đối tượng nhận tiền)",
    mstcus: "VAT Reg. No.\n(MST của đối tượng nhận tiền)",
    citycus: "City\n(Tỉnh của đối tượng nhận tiền)",
    countrycus: "Country\n(Quốc gia của đối tượng nhận tiền)",
    countrygl: "Country GL",
    plant: "Plant\n(Nhà máy)",
    productgroupmot: "Product group",
    producttype: "Product type",
    orderid: "Order ID",
    material: "Material",
  };

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
    fill: { fgColor: { rgb: "005970" } },
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
      const cleanData = Object.fromEntries(
        Object.entries(COLUMN_DATA).filter(([, v]) => v !== undefined),
      );
      const headerKeys = Object.keys(cleanData);
      const headerLabels = Object.values(cleanData);
      const sheet = XLSX.utils.aoa_to_sheet([headerKeys, headerLabels]);
      const range = XLSX.utils.decode_range(sheet["!ref"]);
      const cols = [];
      for (let C = range.s.c; C <= range.e.c; ++C) {
        cols.push({ wch: 25 });
        const keyAddr = XLSX.utils.encode_cell({ c: C, r: 0 });
        const lblAddr = XLSX.utils.encode_cell({ c: C, r: 1 });
        if (!sheet[keyAddr]) sheet[keyAddr] = { v: "" };
        if (!sheet[lblAddr]) sheet[lblAddr] = { v: "" };
        sheet[keyAddr].s = KEY_STYLE;
        sheet[lblAddr].s = LABEL_STYLE;
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
      XLSX.writeFile(wb, "TEMPLATE_FI_DOCUMENT_UPLOAD_" + sStamp + ".xlsx");
    },
  };
});
