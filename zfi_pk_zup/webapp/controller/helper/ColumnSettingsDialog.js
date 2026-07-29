sap.ui.define(
    [
        "sap/m/Dialog",
        "sap/m/Button",
        "sap/m/List",
        "sap/m/StandardListItem",
        "sap/m/SearchField",
        "sap/m/Toolbar",
        "sap/m/ToolbarSpacer",
        "sap/m/CheckBox",
        "sap/m/MessageToast",
        "sap/ui/model/json/JSONModel",
        "sap/ui/model/Filter",
        "sap/ui/model/FilterOperator",
        "./ColumnConfig",
    ],
    function (
        Dialog,
        Button,
        List,
        StandardListItem,
        SearchField,
        Toolbar,
        ToolbarSpacer,
        CheckBox,
        MessageToast,
        JSONModel,
        Filter,
        FilterOperator,
        ColumnConfig
    ) {
        "use strict";

        // Dialog chọn cột tự viết, không phụ thuộc sap.m.p13n
        // (Engine.getInstance không tương thích 1.108).
        // Dùng chung 1 instance dialog cho controller, lazy init.

        const STORAGE_KEY = "zfipkzup.fi.visibleCols";

        function _loadSaved() {
            try {
                const sRaw = window.localStorage.getItem(STORAGE_KEY);
                if (!sRaw) return null;
                const a = JSON.parse(sRaw);
                return Array.isArray(a) && a.length > 0 ? a : null;
            } catch (e) {
                return null;
            }
        }

        function _save(aKeys) {
            try {
                window.localStorage.setItem(STORAGE_KEY, JSON.stringify(aKeys));
            } catch (e) {
                // localStorage bị chặn thì bỏ qua, app vẫn chạy bình thường
            }
        }

        return {
            /**
             * Trả về danh sách key cột visible khi khởi tạo:
             * ưu tiên lựa chọn đã lưu localStorage, không có thì lấy default 13 cột.
             */
            getInitialVisibleKeys: function () {
                return _loadSaved() || ColumnConfig.getDefaultVisibleKeys();
            },

            saveVisibleKeys: function (aKeys) {
                _save(aKeys);
            },

            /**
             * Mở dialog chọn cột.
             * @param {sap.ui.core.mvc.Controller} oController controller Fi
             * @param {function(string[])} fnApply callback nhận mảng key cột visible
             */
            open: function (oController, fnApply) {
                if (!oController._oColDialog) {
                    this._create(oController);
                }
                const oDialog = oController._oColDialog;
                const oModel = oDialog.getModel("colDlg");

                // Đồng bộ trạng thái hiện tại từ bảng vào dialog mỗi lần mở
                const oTable = oController.byId("tableUpload");
                const mVisible = {};
                oTable.getColumns().forEach(function (oCol) {
                    mVisible[oCol.data("colKey")] = oCol.getVisible();
                });
                const aCols = ColumnConfig.getColumns().map(function (c) {
                    return {
                        key: c.key,
                        label: c.label,
                        selected: mVisible[c.key] === true,
                    };
                });
                oModel.setProperty("/cols", aCols);

                // Reset search
                oDialog._oSearch.setValue("");
                oDialog._oList.getBinding("items").filter([]);

                oDialog._fnApply = fnApply;
                oDialog.open();
            },

            _create: function (oController) {
                const oBundle = oController.getView().getModel("i18n").getResourceBundle();
                const oModel = new JSONModel({ cols: [] });
                oModel.setSizeLimit(200); // mặc định 100, ta có 88 cột + dư địa

                const oSearch = new SearchField({
                    width: "100%",
                    placeholder: "Tìm cột...",
                    liveChange: function (oEvt) {
                        const sQuery = oEvt.getParameter("newValue") || "";
                        const oBinding = oList.getBinding("items");
                        if (!sQuery) {
                            oBinding.filter([]);
                            return;
                        }
                        oBinding.filter([
                            new Filter({
                                filters: [
                                    new Filter("label", FilterOperator.Contains, sQuery),
                                    new Filter("key", FilterOperator.Contains, sQuery),
                                ],
                                and: false,
                            }),
                        ]);
                    },
                });

                const oSelectAll = new CheckBox({
                    text: oBundle.getText ? oBundle.getText("SELECT_ALL_FILTERED") || "Chọn tất cả (theo kết quả lọc)" : "Chọn tất cả (theo kết quả lọc)",
                    select: function (oEvt) {
                        const bSel = oEvt.getParameter("selected");
                        // Chỉ tác động lên item đang hiển thị sau filter
                        oList.getItems().forEach(function (oItem) {
                            const oCtx = oItem.getBindingContext("colDlg");
                            if (oCtx) {
                                oModel.setProperty(oCtx.getPath() + "/selected", bSel);
                            }
                        });
                    },
                });

                const oList = new List({
                    mode: "MultiSelect",
                    includeItemInSelection: true,
                    growing: true,
                    growingThreshold: 100,
                    items: {
                        path: "colDlg>/cols",
                        template: new StandardListItem({
                            title: "{colDlg>label}",
                            description: "{colDlg>key}",
                            selected: "{colDlg>selected}",
                        }),
                    },
                });

                const oDialog = new Dialog({
                    title: "Tùy chỉnh cột hiển thị",
                    contentWidth: "30rem",
                    contentHeight: "34rem",
                    draggable: true,
                    resizable: true,
                    subHeader: new Toolbar({ content: [oSearch] }),
                    content: [
                        new Toolbar({ content: [oSelectAll, new ToolbarSpacer()] }),
                        oList,
                    ],
                    buttons: [
                        new Button({
                            text: oBundle.getText("DEFAULT") || "Mặc định",
                            icon: "sap-icon://reset",
                            press: function () {
                                const aDefault = ColumnConfig.getDefaultVisibleKeys();
                                const aCols = oModel.getProperty("/cols");
                                aCols.forEach(function (c) {
                                    c.selected = aDefault.indexOf(c.key) !== -1;
                                });
                                oModel.setProperty("/cols", aCols);
                            },
                        }),
                        new Button({
                            text: oBundle.getText("APPLY") || "Áp dụng",
                            type: "Emphasized",
                            press: function () {
                                const aSelected = oModel
                                    .getProperty("/cols")
                                    .filter(function (c) { return c.selected; })
                                    .map(function (c) { return c.key; });

                                if (aSelected.length === 0) {
                                    MessageToast.show(oBundle.getText("MUST_SELECT_AT_LEAST") || "Phải chọn ít nhất 1 cột");
                                    return;
                                }
                                _save(aSelected);
                                if (oDialog._fnApply) {
                                    oDialog._fnApply(aSelected);
                                }
                                oDialog.close();
                            },
                        }),
                        new Button({
                            text: oBundle.getText("CANCEL") || "Hủy",
                            press: function () { oDialog.close(); },
                        }),
                    ],
                });

                oDialog._oSearch = oSearch;
                oDialog._oList = oList;
                oDialog.setModel(oModel, "colDlg");

                oController.getView().addDependent(oDialog);
                oController._oColDialog = oDialog;
            },

            destroy: function (oController) {
                if (oController._oColDialog) {
                    oController._oColDialog.destroy();
                    oController._oColDialog = null;
                }
            },
        };
    }
);