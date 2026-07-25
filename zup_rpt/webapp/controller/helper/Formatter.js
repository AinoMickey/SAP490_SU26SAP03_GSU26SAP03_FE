sap.ui.define([], function () {
    "use strict";

    return {
        /**
         * Nhận "2026-06-12" (Edm.Date) hoặc "20260612" (dats thô) -> "12/06/2026".
         * Rỗng / "0000-00-00" -> "".
         */
        date: function (v) {
            if (!v) return "";
            const s = String(v);
            if (s.indexOf("0000") === 0) return "";
            if (s.length === 10 && s.charAt(4) === "-") {
                return s.substring(8, 10) + "/" + s.substring(5, 7) + "/" + s.substring(0, 4);
            }
            if (s.length === 8) {
                return s.substring(6, 8) + "/" + s.substring(4, 6) + "/" + s.substring(0, 4);
            }
            return s;
        },

        /** Số lượng: bỏ số 0 thập phân thừa, ngăn cách nghìn kiểu VN */
        qty: function (v) {
            if (v === null || v === undefined || v === "") return "";
            const n = Number(v);
            if (isNaN(n)) return String(v);
            return n.toLocaleString("vi-VN", { maximumFractionDigits: 3 });
        },

        /** Tiền: ngăn cách nghìn, 2 số lẻ nếu có */
        amount: function (v) {
            if (v === null || v === undefined || v === "") return "";
            const n = Number(v);
            if (isNaN(n)) return String(v);
            return n.toLocaleString("vi-VN", { maximumFractionDigits: 2 });
        },

        /** Màu ObjectStatus theo trạng thái lệnh SX */
        ppStatusState: function (s) {
            switch (s) {
                case "Released": return "Information";
                case "Completed": return "Success";
                case "Closed": return "Success";
                case "Created": return "Warning";
                default: return "None";
            }
        },

        criticalityToState: function (iCrit) {
            switch (String(iCrit)) {
                case "3": return "Success";
                case "1": return "Error";
                case "2": return "Warning";
                default: return "None";
            }
        },
                uploadStatusToState: function (s) {
            switch (s) {
                case "S": return "Success";
                case "E": return "Error";
                case "P": return "Warning";
                case "R": return "None";
                default: return "None";
            }
        },
    };
});
