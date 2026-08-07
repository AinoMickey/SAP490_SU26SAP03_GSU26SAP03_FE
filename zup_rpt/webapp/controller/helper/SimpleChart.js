sap.ui.define([], function () {
    "use strict";

    // Chart cột SVG thuần - thay sap.viz (sap.viz kéo RequireJS vào trang,
    // xung đột với SheetJS bundle: bundle thấy global require -> tưởng Node
    // -> require("stream") -> chết module. Vẽ tay thì không đụng ai.

    function _esc(s) {
        return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
    }

    return {
        /**
         * @param {Array<{label:string, values:number[]}>} aData
         * @returns {string} SVG markup (1 root element)
         */
        stackedColumns: function (aData) {
            if (!aData || !aData.length) {
                return '<svg xmlns="http://www.w3.org/2000/svg" width="260" height="60">' +
                       '<text x="12" y="34" font-size="15" fill="#666" font-family="Arial">Không có dữ liệu</text></svg>';
            }
            const iBarW = 80, iGap = 36, iChartH = 260, iTopPad = 28, iBottomPad = 44, iLeftPad = 20;
            const fMax = Math.max.apply(null, aData.map((d) => d.values.reduce((sum, v) => sum + v, 0))) || 1;
            const iW = iLeftPad * 2 + aData.length * (iBarW + iGap) - iGap;
            const iH = iTopPad + iChartH + iBottomPad;
            const aColors = ["#0f6e56", "#854f0b", "#a32d2d"];

            let s = '<svg xmlns="http://www.w3.org/2000/svg" width="' + iW + '" height="' + iH +
                    '" font-family="Arial,Helvetica,sans-serif">';
            s += '<line x1="0" y1="' + (iTopPad + iChartH + 0.5) + '" x2="' + iW + '" y2="' +
                 (iTopPad + iChartH + 0.5) + '" stroke="#ccc" stroke-width="1"/>';

            aData.forEach((d, i) => {
                const x = iLeftPad + i * (iBarW + iGap);
                let y = iTopPad + iChartH;
                d.values.forEach((value, idx) => {
                    const iBarH = Math.max(2, Math.round((value / fMax) * iChartH));
                    y -= iBarH;
                    s += '<rect x="' + x + '" y="' + y + '" width="' + iBarW + '" height="' + iBarH +
                         '" rx="3" fill="' + aColors[idx] + '"/>';
                });
                s += '<text x="' + (x + iBarW / 2) + '" y="' + (iTopPad + iChartH + 18) +
                     '" text-anchor="middle" font-size="12" fill="#666">' + _esc(d.label) + '</text>';
            });
            return s + "</svg>";
        },

        /**
         * @param {Array<{label:string, value:number}>} aData
         * @returns {string} SVG markup (1 root element)
         */
        columns: function (aData) {
            if (!aData || !aData.length) {
                return '<svg xmlns="http://www.w3.org/2000/svg" width="200" height="40">' +
                       '<text x="8" y="24" font-size="13" fill="#666" font-family="Arial">Không có dữ liệu</text></svg>';
            }
            const iBarW = 56, iGap = 24, iChartH = 170, iTopPad = 24, iBottomPad = 28, iLeftPad = 12;
            const fMax = Math.max.apply(null, aData.map((d) => d.value)) || 1;
            const iW = iLeftPad * 2 + aData.length * (iBarW + iGap) - iGap;
            const iH = iTopPad + iChartH + iBottomPad;

            let s = '<svg xmlns="http://www.w3.org/2000/svg" width="' + iW + '" height="' + iH +
                    '" font-family="Arial,Helvetica,sans-serif">';
            // trục đáy
            s += '<line x1="0" y1="' + (iTopPad + iChartH + 0.5) + '" x2="' + iW + '" y2="' +
                 (iTopPad + iChartH + 0.5) + '" stroke="#ccc" stroke-width="1"/>';

            aData.forEach((d, i) => {
                const x = iLeftPad + i * (iBarW + iGap);
                const iBarH = Math.max(2, Math.round((d.value / fMax) * iChartH));
                const y = iTopPad + iChartH - iBarH;
                const sVal = Number(d.value).toLocaleString("vi-VN", { maximumFractionDigits: 2 });
                s += '<rect x="' + x + '" y="' + y + '" width="' + iBarW + '" height="' + iBarH +
                     '" rx="3" fill="#0a6ed1"/>';
                s += '<text x="' + (x + iBarW / 2) + '" y="' + (y - 6) +
                     '" text-anchor="middle" font-size="12" fill="#333">' + _esc(sVal) + '</text>';
                s += '<text x="' + (x + iBarW / 2) + '" y="' + (iTopPad + iChartH + 18) +
                     '" text-anchor="middle" font-size="12" fill="#666">' + _esc(d.label) + '</text>';
            });
            return s + "</svg>";
        },

        groupedColumns: function (aData) {
            if (!aData || !aData.length) {
                return '<svg xmlns="http://www.w3.org/2000/svg" width="200" height="40">' +
                       '<text x="8" y="24" font-size="13" fill="#666" font-family="Arial">Không có dữ liệu</text></svg>';
            }
            const iGroupW = 90, iBarW = 24, iGap = 14, iChartH = 170, iTopPad = 24, iBottomPad = 28, iLeftPad = 12;
            const fMax = Math.max.apply(null, aData.map((d) => Math.max(d.successRate, d.errorRate))) || 1;
            const iW = iLeftPad * 2 + aData.length * (iGroupW + iGap) - iGap;
            const iH = iTopPad + iChartH + iBottomPad;
            const aColors = ["#0f6e56", "#a32d2d"];

            let s = '<svg xmlns="http://www.w3.org/2000/svg" width="' + iW + '" height="' + iH +
                    '" font-family="Arial,Helvetica,sans-serif">';
            s += '<line x1="0" y1="' + (iTopPad + iChartH + 0.5) + '" x2="' + iW + '" y2="' +
                 (iTopPad + iChartH + 0.5) + '" stroke="#ccc" stroke-width="1"/>';

            aData.forEach((d, i) => {
                const xGroup = iLeftPad + i * (iGroupW + iGap);
                const xSuccess = xGroup;
                const xError = xGroup + iBarW + 10;
                const successH = Math.max(2, Math.round((d.successRate / fMax) * iChartH));
                const errorH = Math.max(2, Math.round((d.errorRate / fMax) * iChartH));
                s += '<rect x="' + xSuccess + '" y="' + (iTopPad + iChartH - successH) + '" width="' + iBarW + '" height="' + successH + '" rx="3" fill="' + aColors[0] + '"/>';
                s += '<rect x="' + xError + '" y="' + (iTopPad + iChartH - errorH) + '" width="' + iBarW + '" height="' + errorH + '" rx="3" fill="' + aColors[1] + '"/>';
                s += '<text x="' + (xGroup + iGroupW / 2) + '" y="' + (iTopPad + iChartH + 18) + '" text-anchor="middle" font-size="12" fill="#666">' + _esc(d.label) + '</text>';
            });
            return s + "</svg>";
        },

        donut: function (aData) {
            if (!aData || !aData.length) {
                return '<svg xmlns="http://www.w3.org/2000/svg" width="260" height="220">' +
                       '<text x="12" y="34" font-size="15" fill="#666" font-family="Arial">Không có dữ liệu</text></svg>';
            }
            const iW = 280, iH = 220, iCx = 140, iCy = 110, iR = 72, iRinner = 40;
            const fTotal = aData.reduce((sum, d) => sum + d.value, 0) || 1;
            const aColors = ["#0f6e56", "#854f0b", "#a32d2d"];
            let s = '<svg xmlns="http://www.w3.org/2000/svg" width="' + iW + '" height="' + iH + '" font-family="Arial,Helvetica,sans-serif">';
            let fStart = 0;
            aData.forEach((d, idx) => {
                const fPortion = d.value / fTotal;
                if (fPortion <= 0) return;
                const fEnd = fStart + fPortion;
                const x1 = iCx + iR * Math.cos(2 * Math.PI * fStart - Math.PI / 2);
                const y1 = iCy + iR * Math.sin(2 * Math.PI * fStart - Math.PI / 2);
                const x2 = iCx + iR * Math.cos(2 * Math.PI * fEnd - Math.PI / 2);
                const y2 = iCy + iR * Math.sin(2 * Math.PI * fEnd - Math.PI / 2);
                const iLarge = fPortion > 0.5 ? 1 : 0;
                s += '<path d="M ' + iCx + ' ' + iCy + ' L ' + x1 + ' ' + y1 + ' A ' + iR + ' ' + iR + ' 0 ' + iLarge + ' 1 ' + x2 + ' ' + y2 + ' Z" fill="' + aColors[idx] + '" />';
                fStart = fEnd;
            });
            s += '<circle cx="' + iCx + '" cy="' + iCy + '" r="' + iRinner + '" fill="#fff" />';
            s += '<text x="' + iCx + '" y="' + (iCy - 10) + '" text-anchor="middle" font-size="15" fill="#333">Tổng</text>';
            s += '<text x="' + iCx + '" y="' + (iCy + 14) + '" text-anchor="middle" font-size="18" fill="#111">' + _esc(fTotal.toString()) + '</text>';
            let yLabel = 20;
            aData.forEach((d, idx) => {
                const sPercent = ((d.value / fTotal) * 100).toFixed(0) + "%";
                s += '<rect x="10" y="' + yLabel + '" width="10" height="10" fill="' + aColors[idx] + '" />';
                s += '<text x="26" y="' + (yLabel + 9) + '" font-size="12" fill="#333">' + _esc(d.label + ': ' + sPercent) + '</text>';
                yLabel += 18;
            });
            return s + "</svg>";
        },
    };
});
