sap.ui.define(
    [
        "sap/ui/model/json/JSONModel",
        "sap/m/Dialog",
        "sap/m/Button",
        "sap/m/Text",
        "sap/m/Label",
        "sap/m/Table",
        "sap/m/Column",
        "sap/m/ColumnListItem",
        "sap/m/Panel",
        "sap/m/MessageToast",
        "sap/ui/layout/form/SimpleForm",
    ],
    function (
        JSONModel,
        Dialog,
        Button,
        Text,
        Label,
        Table,
        Column,
        ColumnListItem,
        Panel,
        MessageToast,
        SimpleForm
    ) {
        "use strict";

        // Dialog chi tiết 1 dòng lịch sử upload (dùng chung Fi + Pp).
        // - Form hiển thị field theo cấu hình (field không có giá trị thì bỏ qua)
        // - Nút copy giá trị chính (số chứng từ / số lệnh) vào clipboard
        // - Tùy chọn: bảng chi tiết (bút toán FI) load on-demand từ model V4 khác

        function _fmt(v, sType) {
            if (v === null || v === undefined || v === "") return "";
            const s = String(v);
            if (sType === "date") {
                if (s.indexOf("0000") === 0) return "";
                if (s.length === 10 && s.charAt(4) === "-") {
                    return s.substring(8, 10) + "/" + s.substring(5, 7) + "/" + s.substring(0, 4);
                }
                if (s.length === 8 && /^\d{8}$/.test(s)) {
                    return s.substring(6, 8) + "/" + s.substring(4, 6) + "/" + s.substring(0, 4);
                }
                return s;
            }
            if (sType === "num") {
                const n = Number(v);
                return isNaN(n) ? s : n.toLocaleString("vi-VN", { maximumFractionDigits: 3 });
            }
            return s;
        }

        function _copy(sText, oView) {
            const oBundle = oView && oView.getModel && oView.getModel("i18n")
                ? oView.getModel("i18n").getResourceBundle()
                : null;
            if (!sText) {
                MessageToast.show(oBundle ? oBundle.getText("COPY_NO_VALUE") : "Không có giá trị để copy");
                return;
            }
            if (navigator.clipboard && navigator.clipboard.writeText) {
                navigator.clipboard.writeText(sText).then(
                    () => MessageToast.show(oBundle ? oBundle.getText("COPY_SUCCESS", [sText]) : "Đã copy: " + sText),
                    () => MessageToast.show(oBundle ? oBundle.getText("COPY_FAIL") : "Copy thất bại")
                );
            } else {
                // Fallback môi trường không có Clipboard API
                const oTa = document.createElement("textarea");
                oTa.value = sText;
                document.body.appendChild(oTa);
                oTa.select();
                try {
                    document.execCommand("copy");
                    MessageToast.show(oBundle ? oBundle.getText("COPY_SUCCESS", [sText]) : "Đã copy: " + sText);
                } catch (e) {
                    MessageToast.show(oBundle ? oBundle.getText("COPY_FAIL") : "Copy thất bại");
                }
                document.body.removeChild(oTa);
            }
        }

        return {
            /**
             * Mở dialog chi tiết.
             * @param {object} oCfg
             *   view    : view gọi (để addDependent)
             *   title   : tiêu đề dialog
             *   record  : object dữ liệu dòng lịch sử
             *   fields  : [{key, label, type?: 'date'|'num'}] thứ tự hiển thị
             *   copy    : {label, value} (optional) nút copy giá trị chính
             *   items   : (optional) cấu hình bảng chi tiết load on-demand:
             *             { model, path, filters, sorterParams?, title,
             *               columns: [{key, label, type?}] }
             */
            open: async function (oCfg) {
                // ===== Form chi tiết =====
                const aFormContent = [];
                (oCfg.fields || []).forEach((f) => {
                    const sVal = _fmt(oCfg.record[f.key], f.type);
                    if (sVal === "") return; // field trống thì ẩn cho gọn
                    aFormContent.push(new Label({ text: f.label }));
                    aFormContent.push(new Text({ text: sVal, wrapping: true }));
                });
                const oForm = new SimpleForm({
                    layout: "ResponsiveGridLayout",
                    labelSpanL: 4, labelSpanM: 4, columnsL: 1, columnsM: 1,
                    editable: false,
                    content: aFormContent,
                });

                const aContent = [oForm];

                // ===== Bảng chi tiết (bút toán) on-demand =====
                if (oCfg.items && oCfg.items.model) {
                    try {
                        const oBinding = oCfg.items.model.bindList(
                            oCfg.items.path, null, [],
                            oCfg.items.filters || [],
                            oCfg.items.sorterParams || {}
                        );
                        const aCtx = await oBinding.requestContexts(0, 999);
                        const aItems = aCtx.map((c) => c.getObject());

                        const oItemTable = new Table({
                            columns: oCfg.items.columns.map(
                                (c) => new Column({
                                    header: new Text({ text: c.label }),
                                    hAlign: c.type === "num" ? "End" : "Begin",
                                })
                            ),
                            items: {
                                path: "itm>/rows",
                                template: new ColumnListItem({
                                    cells: oCfg.items.columns.map(
                                        (c) => new Text({
                                            text: {
                                                path: "itm>" + c.key,
                                                formatter: (v) => _fmt(v, c.type),
                                            },
                                        })
                                    ),
                                }),
                            },
                        });
                        oItemTable.setModel(new JSONModel({ rows: aItems }), "itm");

                        aContent.push(new Panel({
                            headerText: (oCfg.items.title || "Chi tiết") + " (" + aItems.length + ")",
                            expandable: false,
                            content: [oItemTable],
                        }));
                    } catch (e) {
                        aContent.push(new Panel({
                            headerText: oCfg.items.title || "Chi tiết",
                            content: [new Text({
                                text: "Không tải được chi tiết: " + (e.message || e),
                            })],
                        }));
                    }
                }

                // ===== Dialog =====
                const oDialog = new Dialog({
                    title: oCfg.title,
                    contentWidth: oCfg.items ? "56rem" : "34rem",
                    contentHeight: "auto",
                    resizable: true,
                    draggable: true,
                    content: aContent,
                    beginButton: oCfg.copy
                        ? new Button({
                              icon: "sap-icon://copy",
                              text: oCfg.copy.label,
                              press: () => _copy(String(oCfg.copy.value || ""), oCfg.view),
                          })
                        : undefined,
                    endButton: new Button({
                        text: (oCfg.view && oCfg.view.getModel("i18n") ? oCfg.view.getModel("i18n").getResourceBundle().getText("CLOSE") : "Đóng"),
                        press: () => oDialog.close(),
                    }),
                    afterClose: () => oDialog.destroy(),
                });
                if (oCfg.view) oCfg.view.addDependent(oDialog);
                oDialog.open();
            },
        };
    }
);