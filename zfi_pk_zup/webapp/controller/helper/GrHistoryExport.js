sap.ui.define([], function () {
    "use strict";

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

    const COLUMNS = [
        { key: "GrNumber", label: "GR Number", width: 14 },
        { key: "Status", label: "Status", width: 8 },
        { key: "MaterialDocument", label: "Mat. Document", width: 16 },
        { key: "Message", label: "Message", width: 40 },
        { key: "CreatedBy", label: "Created By", width: 12 },
        { key: "CreatedAt", label: "Created At", width: 20 },
        { key: "LastChangedAt", label: "Cập nhật lúc", width: 20 },
        { key: "BatchId", label: "Batch ID", width: 24 },
    ];

    function _pad(n) { return String(n).padStart(2, "0"); }
    function _stamp() {
        const d = new Date();
        return d.getFullYear() + _pad(d.getMonth() + 1) + _pad(d.getDate()) +
            "_" + _pad(d.getHours()) + _pad(d.getMinutes()) + _pad(d.getSeconds());
    }

    return {
        export: function (aRows) {
            const aHeader = COLUMNS.map((c) => c.label);
            const aBody = aRows.map((row) => COLUMNS.map((c) => row[c.key] ?? ""));
            const oSheet = XLSX.utils.aoa_to_sheet([aHeader].concat(aBody));
            oSheet["!cols"] = COLUMNS.map((c) => ({ wch: c.width }));
            oSheet["!rows"] = [{ hpx: 22 }];
            for (let c = 0; c < COLUMNS.length; c++) {
                const addr = XLSX.utils.encode_cell({ r: 0, c });
                if (oSheet[addr]) oSheet[addr].s = HEADER_STYLE;
            }
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, oSheet, "Lich su GR");
            XLSX.writeFile(wb, "LICH_SU_GR_" + _stamp() + ".xlsx");
        },
    };
});
