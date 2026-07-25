sap.ui.define(
    ["sap/ui/core/mvc/Controller", "sap/ui/model/json/JSONModel"],
    function (Controller, JSONModel) {
        "use strict";

        return Controller.extend("zuprpt.controller.Main", {
            onInit() {
                this.getView().setModel(
                    new JSONModel({ fiTotal: 0, ppTotal: 0, grTotal: 0 }),
                    "stat"
                );
                this._loadStats();
                this.getOwnerComponent()
                    .getRouter()
                    .getRoute("RouteMain")
                    .attachPatternMatched(this._loadStats, this);
            },

            _loadStats() {
                const oModel = this.getView().getModel();
                if (!oModel) return;
                const oStat = this.getView().getModel("stat");

                const oGrModel = this.getView().getModel("gr");

                this._countEntity(oModel, "/FIUploadReport").then((n) =>
                    oStat.setProperty("/fiTotal", n)
                );
                this._countEntity(oModel, "/PPUploadReport").then((n) =>
                    oStat.setProperty("/ppTotal", n)
                );

                if (oGrModel) {
                    this._countEntity(oGrModel, "/GrUpload").then((n) =>
                        oStat.setProperty("/grTotal", n)
                    );
                }
            },

            /** Đếm bằng $count qua header context, không kéo data */
            _countEntity(oModel, sPath) {
                const oBinding = oModel.bindList(sPath, null, null, null, {
                    $count: true,
                });
                return oBinding
                    .requestContexts(0, 1)
                    .then(() => oBinding.getHeaderContext().requestProperty("$count"))
                    .then((n) => Number(n) || 0)
                    .catch(() => 0);
            },

            onPressFiRpt() {
                this.getOwnerComponent().getRouter().navTo("RouteFiRpt");
            },
            onPressPpRpt() {
                this.getOwnerComponent().getRouter().navTo("RoutePpRpt");
            },

            onPressGrRpt() {
                this.getOwnerComponent().getRouter().navTo("RouteGrRpt");
            },

            onPressKpiRpt() {
                this.getOwnerComponent().getRouter().navTo("RouteKpiRpt");
            },


        });
    }
);
