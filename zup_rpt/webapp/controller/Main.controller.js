sap.ui.define(
    ["sap/ui/core/mvc/Controller"],
    function (Controller) {
        "use strict";

        return Controller.extend("zuprpt.controller.Main", {
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
