sap.ui.define(["./Formatter"], function (Formatter) {
    "use strict";


    // Sheet "Summary": tiêu đề báo cáo, thời điểm xuất, tiêu chí lọc, KPI.
    // Sheet "Data": bảng dữ liệu đang hiển thị, header có style.
    // Yêu cầu XLSX (xlsx.bundle) đã được load global bởi controller gọi.

    const BORDER = {
        top: { style: "thin", color: { rgb: "000000" } },
        bottom: { style: "thin", color: { rgb: "000000" } },
        left: { style: "thin", color: { rgb: "000000" } },
        right: { style: "thin", color: { rgb: "000000" } },
    };
    const HEADER_STYLE = {
        fill: { fgColor: { rgb: "005970" } },
        font: { bold: true, sz: 11, name: "Calibri", color: { rgb: "FFFFFF" } },
        alignment: { horizontal: "center", vertical: "center", wrapText: true },
        border: BORDER,
    };
    const TITLE_STYLE = {
        font: { bold: true, sz: 14, name: "Calibri", color: { rgb: "005970" } },
    };
    const KPI_LABEL_STYLE = {
        fill: { fgColor: { rgb: "EAEAEA" } },
        font: { bold: true, sz: 11, name: "Calibri" },
        border: BORDER,
    };
    const KPI_VALUE_STYLE = {
        font: { sz: 11, name: "Calibri" },
        alignment: { horizontal: "right" },
        border: BORDER,
    };

    function _pad(n) { return String(n).padStart(2, "0"); }

    function _stamp() {
        const d = new Date();
        return (
            d.getFullYear() + _pad(d.getMonth() + 1) + _pad(d.getDate()) +
            "_" + _pad(d.getHours()) + _pad(d.getMinutes()) + _pad(d.getSeconds())
        );
    }

    function _now() {
        const d = new Date();
        return (
            _pad(d.getDate()) + "/" + _pad(d.getMonth() + 1) + "/" + d.getFullYear() +
            " " + _pad(d.getHours()) + ":" + _pad(d.getMinutes())
        );
    }

    return {
        /**
         * @param {object} oCfg
         *   title      : tiêu đề báo cáo (hiện trên Summary)
         *   filePrefix : prefix tên file (sẽ gắn timestamp)
         *   filters    : [{label, value}] tiêu chí lọc đang áp
         *   kpis       : [{label, value}] KPI
         *   columns    : [{key, label, width, type?: 'date'|'qty'}]
         *   rows       : mảng object dữ liệu
         */
        export: function (oCfg) {
            // ===== Sheet Summary =====
            const aSummary = [];
            aSummary.push([oCfg.title]);
            aSummary.push(["Thời điểm xuất", _now()]);
            aSummary.push(["Số dòng dữ liệu", oCfg.rows.length]);
            aSummary.push([]);
            aSummary.push(["TIÊU CHÍ LỌC", ""]);
            (oCfg.filters || []).forEach((f) => aSummary.push([f.label, f.value || "(tất cả)"]));
            aSummary.push([]);
            aSummary.push(["KPI", ""]);
            (oCfg.kpis || []).forEach((k) => aSummary.push([k.label, k.value]));

            const oSumSheet = XLSX.utils.aoa_to_sheet(aSummary);
            oSumSheet["!cols"] = [{ wch: 28 }, { wch: 40 }];
            if (oSumSheet["A1"]) oSumSheet["A1"].s = TITLE_STYLE;
            // Style các dòng label/value từ dòng 2 trở đi
            for (let r = 1; r < aSummary.length; r++) {
                if (!aSummary[r] || aSummary[r].length === 0) continue;
                const a = XLSX.utils.encode_cell({ r: r, c: 0 });
                const b = XLSX.utils.encode_cell({ r: r, c: 1 });
                if (oSumSheet[a]) oSumSheet[a].s = KPI_LABEL_STYLE;
                if (oSumSheet[b]) oSumSheet[b].s = KPI_VALUE_STYLE;
            }

            // ===== Sheet Data =====
            const aHeader = oCfg.columns.map((c) => c.label);
            const aBody = oCfg.rows.map((row) =>
                oCfg.columns.map((c) => {
                    const v = row[c.key];
                    if (c.type === "date") return Formatter.date(v);
                    if (c.type === "qty") return v === null || v === undefined ? "" : Number(v);
                    return v === null || v === undefined ? "" : v;
                })
            );
            const oDataSheet = XLSX.utils.aoa_to_sheet([aHeader].concat(aBody));
            oDataSheet["!cols"] = oCfg.columns.map((c) => ({ wch: c.width || 16 }));
            oDataSheet["!rows"] = [{ hpx: 28 }];
            for (let cIdx = 0; cIdx < oCfg.columns.length; cIdx++) {
                const addr = XLSX.utils.encode_cell({ r: 0, c: cIdx });
                if (oDataSheet[addr]) oDataSheet[addr].s = HEADER_STYLE;
            }

            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, oSumSheet, "Summary");
            XLSX.utils.book_append_sheet(wb, oDataSheet, "Data");
            XLSX.writeFile(wb, oCfg.filePrefix + "_" + _stamp() + ".xlsx");
        },
    };
});
