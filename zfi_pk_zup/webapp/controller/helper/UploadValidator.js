/* eslint-disable */
/**
 * UploadValidator - Chặn lỗi SYNTAX ngay khi upload Excel (client-side),
 * trước khi dữ liệu vào bảng. Nút Check chỉ còn lo validate NGHIỆP VỤ (backend).
 *
 * Viết dạng UMD: chạy được trong UI5 (sap.ui.define) lẫn Node (unit test).
 * Mọi hàm thuần (pure), không đụng UI.
 */
(function (global, factory) {
    if (typeof sap !== "undefined" && sap.ui && sap.ui.define) {
        sap.ui.define([], factory);
    } else if (typeof module !== "undefined" && module.exports) {
        module.exports = factory();
    } else {
        global.UploadValidator = factory();
    }
})(this, function () {
    "use strict";

    // ═══════════════ Cấu hình cột template ═══════════════

    var PP_REQUIRED_COLS = [
        "id_doc", "ordertype", "productionplant", "material", "productionversion",
        "totalqty", "baseunit", "datestart", "dateend", "saleorder", "saleorderitem", "longtext",
    ];
    // FI 88 cột: chỉ bắt buộc tồn tại các cột cốt lõi (user hay xóa cột thừa cho gọn,
    // cho phép — miễn cột cốt lõi còn)
    var FI_REQUIRED_COLS = [
        "id_doc", "documentdate", "postingdate", "documenttype", "companycode",
        "currency", "postingkey", "account", "amountindoumentcurrency",
    ];

    // Alias key template cũ (app cloud PP) -> key chuẩn
    var KEY_ALIASES = {
        startbasicdates: "datestart",
        endbasicdates: "dateend",
        salesorder: "saleorder",
        salesorderitem: "saleorderitem",
        note: "longtext",
    };

    // Posting key: phân vế để check cân bằng chứng từ
    var PK_DEBIT = { "40": 1, "01": 1, "21": 1, "70": 1 };
    var PK_CREDIT = { "50": 1, "11": 1, "31": 1, "75": 1 };

    // ═══════════════ Utils ═══════════════

    function _s(v) {
        return v === undefined || v === null ? "" : String(v).trim();
    }

    /** Chuẩn hóa key: lowercase + trim + alias */
    function normalizeKeys(oRow) {
        var out = {};
        Object.keys(oRow || {}).forEach(function (k) {
            var sKey = String(k).toLowerCase().trim();
            if (KEY_ALIASES[sKey]) sKey = KEY_ALIASES[sKey];
            // trim value: user hay dính khoảng trắng / xuống dòng khi copy-paste
            out[sKey] = oRow[k] === undefined || oRow[k] === null
                ? oRow[k] : String(oRow[k]).trim();
        });
        return out;
    }

    /** Dòng trống = mọi giá trị rỗng */
    function isEmptyRow(r) {
        return Object.keys(r).every(function (k) { return _s(r[k]) === ""; });
    }

    /**
     * Parse ngày: nhận DD/MM/YYYY hoặc DD.MM.YYYY, kiểm tra ngày có thật.
     * @returns {value: 'DD/MM/YYYY', valid: bool}
     */
    function parseDate(v) {
        var s = _s(v).replace(/[.\-]/g, "/");
        if (s === "") return { value: "", valid: false };
        var m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(s);
        if (!m) return { value: s, valid: false };
        var dd = Number(m[1]), mm = Number(m[2]), yyyy = Number(m[3]);
        var d = new Date(yyyy, mm - 1, dd);
        var ok = d.getFullYear() === yyyy && d.getMonth() === mm - 1 && d.getDate() === dd;
        var pad = function (n) { return (n < 10 ? "0" : "") + n; };
        return { value: ok ? pad(dd) + "/" + pad(mm) + "/" + yyyy : s, valid: ok };
    }

    function dateToNum(sDDMMYYYY) {
        // 'DD/MM/YYYY' -> 20260731 để so sánh
        return Number(sDDMMYYYY.substring(6) + sDDMMYYYY.substring(3, 5) + sDDMMYYYY.substring(0, 2));
    }

    /**
     * Parse số chấp nhận cả kiểu VN lẫn quốc tế:
     *  "1234" / "1234.56" / "1234,56" / "1.234,56" / "1,234.56" / "1.234" (VN nghìn)
     * Quy tắc: có cả '.' và ',' -> dấu xuất hiện SAU CÙNG là thập phân.
     * Chỉ có 1 loại dấu, xuất hiện 1 lần và sau nó đúng 3 chữ số -> coi là dấu nghìn.
     * @returns number | NaN
     */
    function parseNumber(v) {
        var s = _s(v);
        if (s === "") return NaN;
        if (!/^-?[\d.,\s]+$/.test(s)) return NaN;
        s = s.replace(/\s/g, "");
        var iDot = s.lastIndexOf("."), iCom = s.lastIndexOf(",");
        if (iDot >= 0 && iCom >= 0) {
            if (iDot > iCom) s = s.replace(/,/g, "");                 // 1,234.56
            else s = s.replace(/\./g, "").replace(",", ".");          // 1.234,56
        } else if (iCom >= 0) {
            var aParts = s.split(",");
            if (aParts.length === 2 && aParts[1].length === 3 && aParts[0].length > 0) {
                s = s.replace(",", "");                               // 1,234 -> nghìn
            } else {
                s = s.replace(/,/g, "").replace(/(\d+)(\d{0})$/, "$1"); // nhiều ',' -> nghìn
                if (aParts.length === 2) s = aParts[0] + "." + aParts[1]; // 12,5 -> thập phân
            }
        } else if (iDot >= 0) {
            var aP = s.split(".");
            if (aP.length === 2 && aP[1].length === 3 && aP[0].length > 0) {
                s = s.replace(".", "");                               // 1.234 -> nghìn (VN)
            } else if (aP.length > 2) {
                s = s.replace(/\./g, "");                             // 1.234.567
            }
        }
        var n = Number(s);
        return isNaN(n) ? NaN : n;
    }

    /** Dòng label của template: nhiều ô chứa '*', '(' hoặc xuống dòng */
    function isLabelRow(oRawRow) {
        var iHits = 0;
        Object.keys(oRawRow || {}).forEach(function (k) {
            var v = oRawRow[k] === undefined || oRawRow[k] === null ? "" : String(oRawRow[k]);
            if (v.indexOf("*") >= 0 || v.indexOf("(") >= 0 || v.indexOf("\n") >= 0) iHits++;
        });
        return iHits >= 2;
    }

    /** Ngày bị Excel đổi thành số serial (vd 45843) */
    function isExcelSerial(sVal) {
        return /^\d{5}(\.0+)?$/.test(sVal);
    }

    function dateErrText(sLabel, sVal) {
        if (isExcelSerial(sVal)) {
            return sLabel + " '" + sVal + "' là số serial - Excel đã tự đổi định dạng ô ngày. " +
                   "Hãy để ô dạng Text và nhập lại DD/MM/YYYY";
        }
        return sLabel + " '" + sVal + "' sai định dạng DD/MM/YYYY hoặc không tồn tại";
    }

    function err(aErrors, iRow, sField, sMsg) {
        aErrors.push({ excelRow: iRow, field: sField, message: sMsg });
    }

    // ═══════════════ API ═══════════════

    return {
        normalizeKeys: normalizeKeys,
        parseDate: parseDate,
        parseNumber: parseNumber,
        isLabelRow: isLabelRow,

        /**
         * Bóc dòng label (row 2 của template) một cách AN TOÀN:
         * chỉ bỏ khi dòng đầu thực sự là label. Nếu user đã xóa row label,
         * dòng đầu là DATA -> giữ nguyên (fix bug delete mù nuốt mất dòng đầu).
         * @returns {rows: Array, startRow: number} startRow = số dòng Excel của data đầu tiên
         */
        stripLabelRow: function (aExcelData) {
            var a = [];
            (aExcelData || []).forEach(function (r) { if (r !== undefined) a.push(r); });
            if (a.length && isLabelRow(a[0])) {
                a.shift();
                return { rows: a, startRow: 3 };
            }
            return { rows: a, startRow: 2 };
        },

        /**
         * Check file có đúng template không (header row 1).
         * @param {string[]} aHeaderKeys mảng key row 1 của sheet
         * @param {'FI'|'PP'} sType
         * @returns {ok, missing[]}
         */
        checkTemplate: function (aHeaderKeys, sType) {
            var aReq = sType === "PP" ? PP_REQUIRED_COLS : FI_REQUIRED_COLS;
            var oHave = {};
            (aHeaderKeys || []).forEach(function (k) {
                var sKey = String(k || "").toLowerCase().trim();
                if (KEY_ALIASES[sKey]) sKey = KEY_ALIASES[sKey];
                oHave[sKey] = 1;
            });
            var aMissing = aReq.filter(function (k) { return !oHave[k]; });
            return { ok: aMissing.length === 0, missing: aMissing };
        },

        /**
         * Chuẩn hóa mảng thô từ sheet_to_row_object_array (sau khi delete [0])
         * thành [{rowNo, r}] — rowNo = số dòng thật trong Excel, bỏ dòng trống.
         * @param {Array} aExcelData mảng sparse, index 1 ~ Excel row 4?? KHÔNG:
         *        index i (bắt đầu 1 sau delete[0]) ~ Excel row i + 2? Quy ước
         *        của cả 2 app: data từ row 3, aExcelData[1] = row 4... nên
         *        rowNo = index + 3 khi duyệt cả index 0 (label đã delete).
         */
        prepareRows: function (aExcelData, iStartRow) {
            var iBase = iStartRow || 3;
            var aOut = [];
            var i = 0;
            (aExcelData || []).forEach(function (raw) {
                if (raw === undefined || raw === null) return;
                var r = normalizeKeys(raw);
                var iRowNo = iBase + i;
                i++;
                if (isEmptyRow(r)) return;
                aOut.push({ rowNo: iRowNo, r: r });
            });
            return aOut;
        },

        // ═══════════════ PP ═══════════════

        /**
         * Validate syntax PP. @param aPrepared từ prepareRows
         * @returns errors[{excelRow, field, message}]
         */
        validatePP: function (aPrepared) {
            var aErrors = [];
            var oSeenId = {};

            aPrepared.forEach(function (o) {
                var r = o.r, n = o.rowNo;

                // required + độ dài
                [["ordertype", 4, "Order Type"],
                 ["productionplant", 4, "Production Plant"],
                 ["material", 40, "Material"],
                 ["productionversion", 4, "Production Version"]].forEach(function (c) {
                    var v = _s(r[c[0]]);
                    if (v === "") err(aErrors, n, c[0], c[2] + " không được để trống");
                    else if (v.length > c[1]) err(aErrors, n, c[0], c[2] + " '" + v + "' vượt quá " + c[1] + " ký tự");
                });

                // totalqty
                var sQty = _s(r.totalqty);
                if (sQty === "") {
                    err(aErrors, n, "totalqty", "Total Quantity không được để trống");
                } else {
                    var fQty = parseNumber(sQty);
                    if (isNaN(fQty)) err(aErrors, n, "totalqty", "Total Quantity '" + sQty + "' không phải là số");
                    else if (fQty <= 0) err(aErrors, n, "totalqty", "Total Quantity phải lớn hơn 0");
                    else if (fQty > 999999999999) err(aErrors, n, "totalqty", "Total Quantity quá lớn");
                }

                // baseunit
                if (_s(r.baseunit).length > 3) err(aErrors, n, "baseunit", "Base Unit vượt quá 3 ký tự");

                // dates
                var oStart = parseDate(r.datestart), oEnd = parseDate(r.dateend);
                if (_s(r.datestart) === "") err(aErrors, n, "datestart", "Start Basic Date không được để trống");
                else if (!oStart.valid) err(aErrors, n, "datestart", dateErrText("Start Basic Date", _s(r.datestart)));
                if (_s(r.dateend) === "") err(aErrors, n, "dateend", "End Basic Date không được để trống");
                else if (!oEnd.valid) err(aErrors, n, "dateend", dateErrText("End Basic Date", _s(r.dateend)));
                if (oStart.valid && oEnd.valid && dateToNum(oEnd.value) < dateToNum(oStart.value)) {
                    err(aErrors, n, "dateend", "End Basic Date (" + oEnd.value + ") không được trước Start Basic Date (" + oStart.value + ")");
                }

                // sales order
                var sSO = _s(r.saleorder), sItem = _s(r.saleorderitem);
                if (sSO !== "" && !/^\d{1,10}$/.test(sSO)) err(aErrors, n, "saleorder", "Sales Order '" + sSO + "' phải là số tối đa 10 chữ số");
                if (sItem !== "" && !/^\d{1,6}$/.test(sItem)) err(aErrors, n, "saleorderitem", "Sales Order Item '" + sItem + "' phải là số tối đa 6 chữ số");
                if (sItem !== "" && sSO === "") err(aErrors, n, "saleorderitem", "Có Sales Order Item nhưng thiếu Sales Order");

                // id_doc trùng
                var sId = _s(r.id_doc);
                if (sId !== "") {
                    if (oSeenId[sId]) err(aErrors, n, "id_doc", "ID '" + sId + "' bị trùng với dòng Excel " + oSeenId[sId]);
                    else oSeenId[sId] = n;
                }
            });
            return aErrors;
        },

        // ═══════════════ FI ═══════════════

        /**
         * Validate syntax FI: rule từng dòng + rule theo chứng từ (id_doc).
         * @returns errors[{excelRow, field, message}]
         */
        validateFI: function (aPrepared) {
            var aErrors = [];
            var oDocs = {}; // id_doc -> {firstRow, header{...}, sumD, sumC, unknownPk}

            aPrepared.forEach(function (o) {
                var r = o.r, n = o.rowNo;

                // ── id_doc: FI bắt buộc (để gộp chứng từ)
                var sId = _s(r.id_doc);
                if (sId === "") err(aErrors, n, "id_doc", "ID (id_doc) không được để trống - dùng để gộp các dòng thành 1 chứng từ");

                // ── header fields
                var sCc = _s(r.companycode);
                if (sCc === "") err(aErrors, n, "companycode", "Company Code không được để trống");
                else if (sCc.length > 4) err(aErrors, n, "companycode", "Company Code '" + sCc + "' vượt quá 4 ký tự");

                var sDt = _s(r.documenttype);
                if (sDt === "") err(aErrors, n, "documenttype", "Document Type không được để trống");
                else if (sDt.length > 2) err(aErrors, n, "documenttype", "Document Type '" + sDt + "' vượt quá 2 ký tự");

                var sCur = _s(r.currency);
                if (sCur === "") err(aErrors, n, "currency", "Currency không được để trống");
                else if (!/^[A-Za-z0-9]{3,5}$/.test(sCur)) err(aErrors, n, "currency", "Currency '" + sCur + "' không hợp lệ (3-5 ký tự, vd VND/USD)");

                var oDocDate = parseDate(r.documentdate), oPstDate = parseDate(r.postingdate);
                if (_s(r.documentdate) === "") err(aErrors, n, "documentdate", "Document Date không được để trống");
                else if (!oDocDate.valid) err(aErrors, n, "documentdate", dateErrText("Document Date", _s(r.documentdate)));
                if (_s(r.postingdate) === "") err(aErrors, n, "postingdate", "Posting Date không được để trống");
                else if (!oPstDate.valid) err(aErrors, n, "postingdate", dateErrText("Posting Date", _s(r.postingdate)));

                // ── item fields
                var sPk = _s(r.postingkey);
                if (sPk === "") err(aErrors, n, "postingkey", "Posting Key không được để trống");
                else if (!/^\d{2}$/.test(sPk)) err(aErrors, n, "postingkey", "Posting Key '" + sPk + "' phải là 2 chữ số (vd 40/50/01/11)");

                var sAcc = _s(r.account);
                if (sAcc === "") err(aErrors, n, "account", "Account không được để trống");
                else if (sAcc.length > 10) err(aErrors, n, "account", "Account '" + sAcc + "' vượt quá 10 ký tự");

                var fAmtDoc = NaN;
                var sAmtDoc = _s(r.amountindoumentcurrency);
                if (sAmtDoc === "") {
                    err(aErrors, n, "amountindoumentcurrency", "Amount Doc Cur không được để trống");
                } else {
                    fAmtDoc = parseNumber(sAmtDoc);
                    if (isNaN(fAmtDoc)) err(aErrors, n, "amountindoumentcurrency", "Amount Doc Cur '" + sAmtDoc + "' không phải là số");
                    else if (fAmtDoc <= 0) err(aErrors, n, "amountindoumentcurrency", "Amount Doc Cur phải lớn hơn 0 (chiều Nợ/Có xác định bằng Posting Key, negative posting dùng cột Negative Posting)");
                }

                // các cột số optional
                [["amountinlocalcurrency", "Amount Local Cur"],
                 ["taxbaseamount", "Tax Base Amount"],
                 ["localtaxbaseamount", "Local Tax Base Amount"],
                 ["exchangerate", "Exchange Rate"],
                 ["quantity", "Quantity"]].forEach(function (c) {
                    var v = _s(r[c[0]]);
                    if (v !== "" && isNaN(parseNumber(v))) err(aErrors, n, c[0], c[1] + " '" + v + "' không phải là số");
                });

                // negativeposting
                var sNeg = _s(r.negativeposting).toUpperCase();
                if (sNeg !== "" && sNeg !== "X") {
                    err(aErrors, n, "negativeposting", "Negative Posting chỉ nhận 'X' hoặc để trống (đang là '" + _s(r.negativeposting) + "')");
                }

                // currency ngoại tệ phải có đường tính local amount
                if (sCur !== "" && sCur.toUpperCase() !== "VND") {
                    var sRate = _s(r.exchangerate);
                    if (_s(r.amountinlocalcurrency) === "" && sRate === "") {
                        err(aErrors, n, "exchangerate", "Currency " + sCur + " (khác VND): phải điền Exchange Rate hoặc Amount Local Cur");
                    }
                    if (sRate !== "" && parseNumber(sRate) === 0) {
                        err(aErrors, n, "exchangerate", "Exchange Rate không được bằng 0");
                    }
                }

                // quantity có thì phải có đơn vị
                if (_s(r.quantity) !== "" && _s(r.unit) === "" && _s(r.baseunit) === "") {
                    err(aErrors, n, "unit", "Có Quantity nhưng thiếu Unit/Base Unit");
                }

                // các cột ngày optional
                [["assetvaluedate", "Asset Value Date"],
                 ["baselinedate", "Baseline Date"],
                 ["valuedate", "Value Date"],
                 ["netduedate", "Net Due Date"]].forEach(function (c) {
                    var v = _s(r[c[0]]);
                    if (v !== "" && !parseDate(v).valid) err(aErrors, n, c[0], c[1] + " '" + v + "' sai định dạng DD/MM/YYYY");
                });

                // ── gom theo chứng từ
                if (sId !== "") {
                    var oDoc = oDocs[sId];
                    if (!oDoc) {
                        oDoc = oDocs[sId] = {
                            firstRow: n, sumD: 0, sumC: 0, unknownPk: false,
                            header: { companycode: sCc, documenttype: sDt, currency: sCur,
                                      documentdate: oDocDate.value, postingdate: oPstDate.value,
                                      negativeposting: sNeg },
                        };
                    } else {
                        // consistency header trong cùng chứng từ
                        [["companycode", sCc, "Company Code"],
                         ["documenttype", sDt, "Document Type"],
                         ["currency", sCur, "Currency"],
                         ["documentdate", oDocDate.value, "Document Date"],
                         ["postingdate", oPstDate.value, "Posting Date"],
                         ["negativeposting", sNeg, "Negative Posting"]].forEach(function (c) {
                            if (oDoc.header[c[0]] !== c[1]) {
                                err(aErrors, n, c[0],
                                    "ID " + sId + ": " + c[2] + " '" + c[1] + "' khác dòng đầu của chứng từ (dòng " +
                                    oDoc.firstRow + " là '" + oDoc.header[c[0]] + "') - các dòng cùng ID phải cùng thông tin header");
                            }
                        });
                    }
                    if (!isNaN(fAmtDoc)) {
                        if (PK_DEBIT[sPk]) oDoc.sumD += fAmtDoc;
                        else if (PK_CREDIT[sPk]) oDoc.sumC += fAmtDoc;
                        else oDoc.unknownPk = true; // PK lạ -> bỏ check cân bằng cho doc này
                    }
                }
            });

            // ── cân bằng Nợ/Có từng chứng từ
            Object.keys(oDocs).forEach(function (sId) {
                var d = oDocs[sId];
                if (d.unknownPk) return;
                if (Math.abs(d.sumD - d.sumC) > 0.01) {
                    err(aErrors, d.firstRow, "amountindoumentcurrency",
                        "ID " + sId + ": chứng từ không cân bằng - tổng Nợ " + d.sumD +
                        " ≠ tổng Có " + d.sumC + " (theo Amount Doc Cur)");
                }
            });

            return aErrors;
        },
    };
});