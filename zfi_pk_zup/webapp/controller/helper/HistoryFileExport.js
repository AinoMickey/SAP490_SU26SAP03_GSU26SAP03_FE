sap.ui.define(
  [
    "sap/ui/model/Filter",
    "sap/ui/model/FilterOperator",
    "zfipkzup/controller/helper/ExcelTemplate",
    "zfipkzup/controller/helper/PpExcelTemplate",
    "zfipkzup/controller/helper/GrExcelTemplate",
  ],
  function (Filter, FilterOperator, ExcelTemplate, PpExcelTemplate, GrExcelTemplate) {
    "use strict";

    const FI_COLS = [
      { k: "id_doc", p: "id_doc" },
      { k: "documentdate", p: "Documentdate", t: "date" },
      { k: "postingdate", p: "Postingdate", t: "date" },
      { k: "documenttype", p: "Documenttype" },
      { k: "companycode", p: "Companycode" },
      { k: "currency", p: "Currency" },
      // Tỷ giá thật nằm ở dòng item; header exchangerate không bao giờ được ghi
      { k: "exchangerate", p: "ItemExchangerate", t: "num" },
      { k: "headertext", p: "Headertext" },
      { k: "referencedoc", p: "Referencedoc" },
      { k: "headerref1", p: "Headerref1" },
      { k: "negativeposting", p: "negativeposting" },
      { k: "postingkey", p: "postingkey" },
      { k: "account", p: "account" },
      { k: "mainassetnumber", p: "mainassetnumber", z0: true },
      { k: "subassetnumber", p: "subassetnumber", z0: true },
      { k: "specialglaccount", p: "specialglaccount" },
      { k: "assettransactiontype", p: "assettransactiontype" },
      { k: "amountindoumentcurrency", p: "amountindoumentcurrency", t: "num", z: true },
      { k: "amountinlocalcurrency", p: "amountinlocalcurrency", t: "num", z: true },
      { k: "taxbaseamount", p: "taxbaseamount", t: "num" },
      { k: "localtaxbaseamount", p: "localtaxbaseamount", t: "num" },
      { k: "assignment", p: "assignment" },
      { k: "costcenter", p: "costcenter", z0: true },
      { k: "profitcenter", p: "profitcenter", z0: true },
      { k: "internalorder", p: "internalorder", z0: true },
      { k: "wbselement", p: "wbselement", z0: true },
      { k: "businessarea", p: "businessarea" },
      { k: "assetvaluedate", p: "assetvaluedate", t: "date" },
      { k: "itemtext", p: "itemtext" },
      { k: "longtext", p: "longtext" },
      { k: "overrideglaccount", p: "overrideglaccount", z0: true },
      { k: "taxcode", p: "taxcode" },
      { k: "segment", p: "segment" },
      { k: "paymentterms", p: "paymentterms" },
      { k: "paymentblockreason", p: "paymentblockreason" },
      { k: "paymentmethod", p: "paymentmethod" },
      { k: "baselinedate", p: "baselinedate", t: "date" },
      { k: "valuedate", p: "valuedate", t: "date" },
      { k: "netduedate", p: "netduedate", t: "date" },
      { k: "contractnumber", p: "contractnumber" },
      { k: "contracttype", p: "contracttype" },
      { k: "housebank", p: "housebank" },
      { k: "bankaccountid", p: "bankaccountid" },
      { k: "invoicerefnum", p: "invoicerefnum", z0: true },
      { k: "invoicereffiscalyear", p: "invoicefiscalyear", z0: true },
      { k: "invoicereflineitem", p: "invoicereflineitem", z0: true },
      { k: "purchasingno", p: "purchasingno", z0: true },
      { k: "purchasingitem", p: "purchasingitem", z0: true },
      { k: "saleorder", p: "saleorder", z0: true },
      { k: "saleorderitem", p: "saleorderitem", z0: true },
      { k: "customer", p: "customer", z0: true },
      { k: "cusgroup", p: "cusgroup" },
      { k: "division", p: "division" },
      { k: "distributionchannel", p: "distributionchannel" },
      { k: "salesorganization", p: "salesorganization" },
      { k: "salesoffice", p: "salesoffice" },
      { k: "salesemployee", p: "salesemployee" },
      { k: "salesgroup", p: "salesgroup" },
      { k: "materialgroup", p: "materialgroup" },
      { k: "product", p: "product", z0: true },
      { k: "unit", p: "unit" },
      { k: "baseunit", p: "baseunit" },
      { k: "quantity", p: "quantity", t: "num" },
      { k: "alternativepayee", p: "alternativepayee", z0: true },
      { k: "name1", p: "name1" },
      { k: "name2", p: "name2" },
      { k: "name3", p: "name3" },
      { k: "name4", p: "name4" },
      { k: "mst", p: "mst" },
      { k: "city", p: "city" },
      { k: "country", p: "country" },
      { k: "vatregno", p: "vatregno" },
      { k: "ref1", p: "ref1" },
      { k: "ref2", p: "ref2" },
      { k: "ref3", p: "ref3" },
      { k: "namecus1", p: "namecus1" },
      { k: "namecus2", p: "namecus2" },
      { k: "namecus3", p: "namecus3" },
      { k: "namecus4", p: "namecus4" },
      { k: "mstcus", p: "mstcus" },
      { k: "citycus", p: "citycus" },
      { k: "countrycus", p: "countrycus" },
      { k: "countrygl", p: "countrygl" },
      { k: "plant", p: "plant" },
      { k: "productgroupmot", p: null },
      { k: "producttype", p: null },
      { k: "orderid", p: "orderid", z0: true },
      { k: "material", p: "material", z0: true },
    ];

    // Cột tham chiếu thêm vào cuối file FI.
    // Cột này KHÔNG thuộc template nên UploadValidator bỏ qua khi upload lại,
    // vì vậy giữ lại được mà không cản trở việc post lại.
    const FI_REF_COLS = [
      {
        k: "accountingdocument",
        p: "accountingdocument",
        label: "Accounting Document\n(Số CT đã post - cột tham chiếu, không thuộc template)",
      },
    ];

    const PP_COLS = [
      { k: "id_doc", p: "IdDoc" },
      { k: "ordertype", p: "ProductionOrderType" },
      { k: "productionplant", p: "ProductionPlant" },
      { k: "material", p: "Material", z0: true },
      { k: "productionversion", p: "ProductionVersion" },
      { k: "totalqty", p: "TotalQty", t: "num", z: true },
      { k: "baseunit", p: "BaseUnit" },
      { k: "datestart", p: "StartDate", t: "date" },
      { k: "dateend", p: "EndDate", t: "date" },
      // Lệnh MTS: log lưu 0, không lưu rỗng -> z0 để file dựng lại để TRỐNG.
      // Không có z0 thì validator chặn "có SO Item nhưng thiếu SO".
      { k: "saleorder", p: "SalesOrder", z0: true },
      { k: "saleorderitem", p: "SalesOrderItem", z0: true },
      // ZPP_TB_ZUPLSX không có cột longtext -> không dựng lại được
      { k: "longtext", p: null },
    ];

    // PP KHÔNG có cột tham chiếu.
    // Số lệnh đã tạo được ghi ở sheet Info thay vì cột Data, để file dựng lại
    // là template sạch và upload lại được ngay.
    const PP_REF_COLS = [];

    const GR_COLS = [
      { k: "gr_number", p: "GrNumber" },
      { k: "document_date", p: "DocumentDate", t: "date" },
      { k: "movement_type", p: "MovementType" },
      { k: "po_number", p: "PoNumber", z0: true },
      { k: "po_item", p: "PoItem", z0: true },
      { k: "receive_qty", p: "ReceiveQty", t: "num", z: true },
      { k: "unit", p: "Unit" },
      { k: "storage_location", p: "StorageLocation" },
    ];

    const GR_REF_COLS = [
      {
        k: "materialdocument",
        p: "MaterialDocument",
        label: "Material Document\n(Chứng từ đã post - cột tham chiếu, không thuộc template)",
      },
      {
        k: "materialdocumentyear",
        p: "MaterialDocumentYear",
        label: "Năm chứng từ\n(cột tham chiếu, không thuộc template)",
      },
    ];

    const MISSING_GR = GR_COLS.filter((c) => !c.p).map((c) => c.k);
    const MISSING_FI = FI_COLS.filter((c) => !c.p).map((c) => c.k);
    const MISSING_PP = PP_COLS.filter((c) => !c.p).map((c) => c.k);

    // ═════════════ Style (bám đúng ExcelTemplate/PpExcelTemplate) ═════════════

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
    function labelStyle(sRgb) {
      return {
        fill: { fgColor: { rgb: sRgb } },
        font: { bold: true, sz: 11, name: "Calibri", color: { rgb: "FFFFFF" } },
        alignment: { horizontal: "center", vertical: "center", wrapText: true },
        border: BORDER,
      };
    }

    // ═════════════ Utils ═════════════

    function _s(v) {
      return v === undefined || v === null ? "" : String(v);
    }

    /**
     * Cột ALPHA/NUMC không bắt buộc: log lưu 0 hoặc 0000000000 thay vì rỗng.
     * Ghi số 0 ra file làm validator hiểu là "có giá trị" và chặn file.
     * Trả về chuỗi rỗng khi mọi ký tự đều là số 0.
     */
    function _stripZeros(s) {
      const t = _s(s).trim();
      return /^0+$/.test(t) ? "" : t;
    }

    /** Edm.Date 'YYYY-MM-DD' hoặc NUMC 'YYYYMMDD' -> 'DD/MM/YYYY'; rỗng/0 -> '' */
    function _fmtDate(v) {
      const s = _s(v).trim();
      if (s === "" || /^0+$/.test(s.replace(/[-/]/g, ""))) return "";
      if (s.length === 10 && s.charAt(4) === "-") {
        return s.substring(8, 10) + "/" + s.substring(5, 7) + "/" + s.substring(0, 4);
      }
      if (/^\d{8}$/.test(s)) {
        return s.substring(6, 8) + "/" + s.substring(4, 6) + "/" + s.substring(0, 4);
      }
      return s;
    }

    /**
     * '1000.00' -> '1000', '23.50000' -> '23.5', '0.000' -> '' (trừ khi bKeepZero).
     * Cắt số 0 thừa bằng CHUỖI, không qua Number(), để không mất chính xác với
     * CURR(23,2) hay DEC(13,5) khi giá trị lớn.
     */
    function _fmtNum(v, bKeepZero) {
      let s = _s(v).trim();
      if (s === "") return "";
      if (!/^-?\d+(\.\d+)?$/.test(s)) return s; // dạng lạ (exponent...) -> giữ nguyên
      if (s.indexOf(".") >= 0) {
        s = s.replace(/0+$/, "").replace(/\.$/, "");
      }
      if (/^-?0*$/.test(s.replace(".", "")) && !bKeepZero) return "";
      return s === "-0" ? "0" : s;
    }

    function _cell(oRow, oCol) {
      if (!oCol.p) return "";
      const v = oRow[oCol.p];
      if (oCol.t === "date") return _fmtDate(v);
      if (oCol.t === "num") return _fmtNum(v, oCol.z);
      const s = _s(v).trim();
      return oCol.z0 ? _stripZeros(s) : s;
    }

    function _stamp() {
      const d = new Date();
      const p = (n) => String(n).padStart(2, "0");
      return (
        d.getFullYear() +
        p(d.getMonth() + 1) +
        p(d.getDate()) +
        "_" +
        p(d.getHours()) +
        p(d.getMinutes()) +
        p(d.getSeconds())
      );
    }

    function _outName(sFilename) {
      const sBase = _s(sFilename).replace(/\.(xlsx|xlsm|xls|csv)$/i, "") || "UPLOAD";
      return sBase + "_REBUILD_" + _stamp() + ".xlsx";
    }

    /** Label lấy từ template gốc; template thiếu key thì fallback = chính key */
    function _labels(oTemplateModule, aRefCols) {
      let oData = {};
      try {
        if (oTemplateModule && typeof oTemplateModule.getColumnData === "function") {
          oData = oTemplateModule.getColumnData() || {};
        }
      } catch (e) {
        oData = {};
      }
      const oOut = Object.assign({}, oData);
      (aRefCols || []).forEach((c) => {
        oOut[c.k] = c.label || c.k;
      });
      return oOut;
    }

    /**
     * Đọc hết dữ liệu qua V4 list binding, phân trang.
     * Thử với $select tường minh trước (autoExpandSelect - bài học #2);
     * nếu service chưa publish đủ field thì fallback bỏ $select.
     */
    async function _readAll(oModel, sPath, aFilters, sOrderBy, aProps) {
      const PAGE = 1000;
      const MAX = 20000;

      async function run(bWithSelect) {
        const oParams = { $orderby: sOrderBy };
        if (bWithSelect && aProps && aProps.length) {
          oParams.$select = aProps.join(",");
        }
        const oBinding = oModel.bindList(sPath, null, [], aFilters, oParams);
        const aOut = [];
        let iStart = 0;
        for (;;) {
          const aCtx = await oBinding.requestContexts(iStart, PAGE);
          aCtx.forEach((c) => aOut.push(c.getObject()));
          if (aCtx.length < PAGE || aOut.length >= MAX) break;
          iStart += PAGE;
        }
        return aOut;
      }

      try {
        return await run(true);
      } catch (e) {
        // $select có field chưa tồn tại trong metadata đã publish -> thử lại không $select
        return await run(false);
      }
    }

    /** Dựng workbook: sheet Data (2 dòng header + data) + sheet Info */
    function _buildWorkbook(aCols, oLabels, aRows, sHeaderRgb, oInfo) {
      if (typeof XLSX === "undefined") {
        throw new Error("Chưa nạp thư viện XLSX (xlsx.bundle) trong controller gọi.");
      }

      const aKeys = aCols.map((c) => c.k);
      const aLbls = aCols.map((c) => oLabels[c.k] || c.k);
      const aoa = [aKeys, aLbls];
      aRows.forEach((r) => aoa.push(aCols.map((c) => _cell(r, c))));

      const oSheet = XLSX.utils.aoa_to_sheet(aoa);
      const oRange = XLSX.utils.decode_range(oSheet["!ref"]);
      const aColW = [];

      for (let C = oRange.s.c; C <= oRange.e.c; ++C) {
        aColW.push({ wch: 25 });
        for (let R = oRange.s.r; R <= oRange.e.r; ++R) {
          const sAddr = XLSX.utils.encode_cell({ r: R, c: C });
          if (!oSheet[sAddr]) oSheet[sAddr] = { t: "s", v: "" };
          oSheet[sAddr].t = "s";
          oSheet[sAddr].z = "@"; // ép Text, tránh Excel tự đổi ngày/số
        }
        oSheet[XLSX.utils.encode_cell({ c: C, r: 0 })].s = KEY_STYLE;
        oSheet[XLSX.utils.encode_cell({ c: C, r: 1 })].s = labelStyle(sHeaderRgb);
      }
      oSheet["!cols"] = aColW;
      oSheet["!rows"] = [{ hpx: 20 }, { hpx: 50 }];

      const oInfoSheet = XLSX.utils.aoa_to_sheet(
        Object.keys(oInfo).map((k) => [k, _s(oInfo[k])]),
      );
      oInfoSheet["!cols"] = [{ wch: 28 }, { wch: 80 }];

      const oWb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(oWb, oSheet, "Data");
      XLSX.utils.book_append_sheet(oWb, oInfoSheet, "Info");
      return oWb;
    }

    return {
      getMissingColumns: function (sType) {
        if (sType === "PP") return MISSING_PP.slice();
        if (sType === "GR") return MISSING_GR.slice();
        return MISSING_FI.slice();
      },

      /**
       * Dựng lại file FI đã post.
       * @param {sap.ui.model.odata.v4.ODataModel} oModel model mặc định (service zfi_ui_zup_fidoc_o4)
       * @param {string} sFilename tên file trong log
       * @param {object} [oOpt] { withRefs: true/false } - kèm cột accountingdocument
       * @returns {Promise<{rows:number, file:string, missing:Array}>}
       */
      exportFi: async function (oModel, sFilename, oOpt) {
        if (!oModel) throw new Error("Thiếu model service FI (zfi_ui_zup_fidoc_o4).");
        if (!sFilename) throw new Error("Thiếu filename.");
        const bRefs = !oOpt || oOpt.withRefs !== false;

        const aCols = bRefs ? FI_COLS.concat(FI_REF_COLS) : FI_COLS.slice();
        const aProps = aCols
          .filter((c) => c.p)
          .map((c) => c.p)
          .concat(["filename", "id_line"]);

        const aRows = await _readAll(
          oModel,
          "/UploadHistoryItem",
          [new Filter("filename", FilterOperator.EQ, sFilename)],
          "id_doc,id_line",
          Array.from(new Set(aProps)),
        );

        if (!aRows.length) {
          throw new Error("Không tìm thấy dòng chi tiết nào của file " + sFilename);
        }

        const sOut = _outName(sFilename);
        const oWb = _buildWorkbook(
          aCols,
          _labels(ExcelTemplate, FI_REF_COLS),
          aRows,
          "005970",
          {
            "Loai upload": "FI - Chung tu ke toan",
            "File goc": sFilename,
            "So dong": aRows.length,
            "Ngay dung lai": new Date().toISOString(),
            "Nguon du lieu": "Log ZFI_TB_UPLOAD / ZFI_TB_UPLOAD_I (du lieu da post)",
            "Cot khong co trong log":
              MISSING_FI.join(", ") +
              " (khong khai trong ts_doc_item_request nen bi bo khi deserialise)",
            "Upload lai": "Doi ten file roi upload. Cot accountingdocument khong thuoc " +
              "template nen validator bo qua, khong can xoa.",
            "Luu y": "Day la file DUNG LAI tu log, khong phai file goc.",
          },
        );

        XLSX.writeFile(oWb, sOut);
        return { rows: aRows.length, file: sOut, missing: MISSING_FI };
      },

      /**
       * Dựng lại file PP đã post.
       * File dựng lại là TEMPLATE SẠCH: không còn cột productionorder, các cột
       * sales order để trống khi lệnh là MTS. Đổi tên file là upload lại được.
       * @param {sap.ui.model.odata.v4.ODataModel} oModel model "pp" (service zpp_ui_zuplsx_o4)
       * @param {string} sFilename tên file trong log
       * @param {object} [oOpt] giữ để tương thích chữ ký cũ, không còn tác dụng
       */
      exportPp: async function (oModel, sFilename, oOpt) {
        if (!oModel) throw new Error("Thiếu model 'pp' (zpp_ui_zuplsx_o4).");
        if (!sFilename) {
          throw new Error(
            "Dòng lịch sử này không có Filename (lệnh tạo trước khi log lưu filename).",
          );
        }

        const aCols = PP_COLS.slice();
        // ProductionOrder vẫn được đọc để ghi vào sheet Info, dù không còn là cột Data
        const aProps = aCols
          .filter((c) => c.p)
          .map((c) => c.p)
          .concat(["Filename", "ProductionOrder"]);

        const aRows = await _readAll(
          oModel,
          "/UploadHistory",
          [new Filter("Filename", FilterOperator.EQ, sFilename)],
          "IdDoc",
          Array.from(new Set(aProps)),
        );

        if (!aRows.length) {
          throw new Error("Không tìm thấy lệnh nào thuộc file " + sFilename);
        }

        const sOrders = aRows
          .map((r) => _s(r.ProductionOrder).trim())
          .filter((s) => s !== "" && !/^0+$/.test(s))
          .join(", ");

        const bNoUnit = aRows.every((r) => _s(r.BaseUnit).trim() === "");

        const sOut = _outName(sFilename);
        const oWb = _buildWorkbook(
          aCols,
          _labels(PpExcelTemplate, PP_REF_COLS),
          aRows,
          "1B5E20",
          {
            "Loai upload": "PP - Lenh san xuat",
            "File goc": sFilename,
            "So dong": aRows.length,
            "Ngay dung lai": new Date().toISOString(),
            "Nguon du lieu": "Log ZPP_TB_ZUPLSX (chi cac dong da tao lenh thanh cong)",
            "Lenh da tao": sOrders || "(khong co)",
            "Upload lai":
              "File nay la template sach: doi ten file roi upload va Post la tao " +
              "duoc lenh moi. Khong can xoa cot nao.",
            "Cot longtext":
              "ZPP_TB_ZUPLSX khong co cot longtext nen ghi chu khong dung lai duoc. " +
              "Dien lai neu can.",
            "Cot baseunit": bNoUnit
              ? "Trong tren moi dong. Du lieu post TRUOC ban va ZPP_CL_ZUPLSX_VALIDATOR " +
                "(nhanh ELSE lay MARA-MEINS) khong luu don vi tinh. De trong van post lai duoc: " +
                "SAP tu lay theo material."
              : "Lay tu log.",
            "Luu y": "Day la file DUNG LAI tu log, khong phai file goc.",
          },
        );

        XLSX.writeFile(oWb, sOut);
        return { rows: aRows.length, file: sOut, missing: MISSING_PP, orders: sOrders };
      },

      /**
       * Dựng lại file GR của một lần upload. Mỗi lần gọi uploadExcel sinh 1 BatchId
       * dùng chung cho mọi GR trong file, nên BatchId chính là định danh của file.
       * Document Date / Movement Type nằm ở header, các cột PO nằm ở item —
       * ghép lại thành 1 dòng Excel đúng như template.
       * @param {sap.ui.model.odata.v4.ODataModel} oModel model "gr" (zmm_ui_pogr_o4)
       * @param {string} sBatchId batch của lần upload cần dựng lại
       * @param {object} [oOpt] { withRefs: true/false } - kèm cột material document
       */
      exportGr: async function (oModel, sBatchId, oOpt) {
        if (!oModel) throw new Error("Thiếu model 'gr' (zmm_ui_pogr_o4).");
        if (!sBatchId) throw new Error("Dòng lịch sử này không có Batch ID.");
        const bRefs = !oOpt || oOpt.withRefs !== false;

        const sBatch = String(sBatchId).trim();
        const aHdrProps = ["GrNumber", "DocumentDate", "MovementType", "BatchId"];
        let aHeaders = await _readAll(
          oModel,
          "/GrUpload",
          [new Filter("BatchId", FilterOperator.EQ, sBatch)],
          "GrNumber",
          aHdrProps,
        );
        if (!aHeaders.length) {
          aHeaders = await _readAll(
            oModel,
            "/GrUpload",
            [new Filter("BatchId", FilterOperator.StartsWith, sBatch)],
            "GrNumber",
            aHdrProps,
          );
        }

        if (!aHeaders.length) {
          throw new Error("Không tìm thấy GR nào thuộc batch " + sBatch);
        }

        const oHdrByGr = {};
        aHeaders.forEach((h) => {
          oHdrByGr[h.GrNumber] = h;
        });

        const aItems = await _readAll(
          oModel,
          "/GrItem",
          [
            new Filter({
              filters: aHeaders.map(
                (h) => new Filter("GrNumber", FilterOperator.EQ, h.GrNumber),
              ),
              and: false,
            }),
          ],
          "GrNumber,Item",
          [
            "GrNumber", "Item", "PoNumber", "PoItem", "ReceiveQty",
            "Unit", "StorageLocation", "MaterialDocument", "MaterialDocumentYear",
          ],
        );
        if (!aItems.length) {
          throw new Error("Batch " + sBatchId + " không có dòng item nào.");
        }

        // 1 dòng Excel = 1 item, kèm field header của GR chứa nó
        const aRows = aItems.map((it) =>
          Object.assign({}, oHdrByGr[it.GrNumber] || {}, it),
        );

        const aCols = bRefs ? GR_COLS.concat(GR_REF_COLS) : GR_COLS.slice();
        const sOutName = _outName("GR_" + sBatch.slice(-8));

        const oWb = _buildWorkbook(
          aCols,
          _labels(GrExcelTemplate, GR_REF_COLS),
          aRows,
          "8A3324",
          {
            "Loai upload": "GR - Phieu nhap kho",
            "Batch ID": sBatchId,
            "So GR trong batch": aHeaders.length,
            "So dong": aRows.length,
            "Ngay dung lai": new Date().toISOString(),
            "Nguon du lieu": "Staging ZMM_TB_GR_H / ZMM_TB_GR_I",
            "Cot khong co trong log": MISSING_GR.join(", ") || "(khong co)",
            "Upload lai":
              "Hai cot tham chieu material document khong thuoc template nen " +
              "validator bo qua, khong can xoa.",
            "Luu y": "Day la file DUNG LAI tu staging, khong phai file goc.",
          },
        );

        XLSX.writeFile(oWb, sOutName);
        return { rows: aRows.length, file: sOutName, missing: MISSING_GR };
      },
    };
  },
);