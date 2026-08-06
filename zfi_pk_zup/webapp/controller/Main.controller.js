sap.ui.define(
    ["sap/ui/core/mvc/Controller", "sap/ui/model/json/JSONModel"],
    function (Controller, JSONModel) {
        "use strict";

        return Controller.extend("zfipkzup.controller.Main", {
            onInit() {
                // Model thống kê
                this.getView().setModel(
                    new JSONModel({ total: 0, lastFile: "—", lastDate: "—" }),
                    "stat"
                );
                this._loadStats();

                // Cập nhật lại stats mỗi khi quay về màn Main
                const oRouter = this.getOwnerComponent().getRouter();
                oRouter.getRoute("RouteMain").attachPatternMatched(this._loadStats, this);
            },

            _loadStats() {
                const oDefaultModel = this.getView().getModel(); // V4 default
                if (!oDefaultModel) return;
                const oPpModel = this.getView().getModel("pp");
                const oGrModel = this.getView().getModel("gr");
                const oStat = this.getView().getModel("stat");

                const aSources = [
                    {
                        model: oDefaultModel,
                        path: "/UploadHistory",
                        dateField: "pst_date",
                        fileField: "filename",
                    },
                    {
                        model: oPpModel,
                        path: "/UploadHistory",
                        dateField: "StartDate",
                        fileField: "Filename",
                    },
                    {
                        model: oGrModel,
                        path: "/GrUpload",
                        dateField: "CreatedAt",
                        fileField: "Filename",
                    },
                ];

                Promise.all(
                    aSources.map((oSource) => {
                        if (!oSource.model) {
                            return Promise.resolve({ count: 0, latest: null });
                        }
                        return this._loadSourceStats(oSource);
                    })
                )
                    .then((aStats) => {
                        const iTotal = aStats.reduce((sum, o) => sum + (o.count || 0), 0);
                        const aLatest = aStats
                            .map((o) => o.latest)
                            .filter(Boolean)
                            .sort((a, b) => this._compareDateDesc(a.date, b.date));

                        oStat.setProperty("/total", iTotal);
                        if (aLatest.length > 0) {
                            oStat.setProperty("/lastFile", aLatest[0].file || "—");
                            oStat.setProperty(
                                "/lastDate",
                                aLatest[0].date ? this._fmtDate(aLatest[0].date) : "—"
                            );
                        } else {
                            oStat.setProperty("/lastFile", "Chưa có");
                            oStat.setProperty("/lastDate", "—");
                        }
                    })
                    .catch(() => {
                        oStat.setProperty("/total", 0);
                        oStat.setProperty("/lastFile", "Chưa có");
                        oStat.setProperty("/lastDate", "—");
                    });
            },

            async _loadSourceStats(oSource) {
                const oBinding = oSource.model.bindList(
                    oSource.path,
                    null,
                    [],
                    [],
                    { $orderby: `${oSource.dateField} desc`, $count: true }
                );

                try {
                    const aContexts = await oBinding.requestContexts(0, 1);
                    const iCount =
                        Number(
                            await oBinding.getHeaderContext().requestProperty("$count")
                        ) || 0;
                    const oLatest = aContexts.length ? aContexts[0].getObject() : null;
                    return {
                        count: iCount,
                        latest: oLatest
                            ? {
                                  date: oLatest[oSource.dateField],
                                  file: oLatest[oSource.fileField] || "—",
                              }
                            : null,
                    };
                } catch (e) {
                    return { count: 0, latest: null };
                }
            },

            _compareDateDesc(sA, sB) {
                const dA = new Date(sA).getTime() || 0;
                const dB = new Date(sB).getTime() || 0;
                return dB - dA;
            },

            _fmtDate(sDate) {
                if (typeof sDate === "string" && sDate.length >= 10) {
                    return sDate.substring(8, 10) + "/" + sDate.substring(5, 7);
                }
                return String(sDate);
            },

            onPressFI() {
                this.getOwnerComponent().getRouter().navTo("RouteFi");
            },

            onPressPP() {
                this.getOwnerComponent().getRouter().navTo("RoutePp");
            },
            
            onPressGR() {
                this.getOwnerComponent().getRouter().navTo("RouteGr");
            },

        });
    }
);