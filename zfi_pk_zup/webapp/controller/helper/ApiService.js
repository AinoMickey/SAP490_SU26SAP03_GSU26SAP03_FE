sap.ui.define(
    ["sap/ui/model/Filter", "sap/ui/model/FilterOperator"],
    function (Filter, FilterOperator) {
        "use strict";

        const ACTION_FQN =
            "com.sap.gateway.srvd.zfi_ui_zup_fidoc.v0001.uploadFromExcel";

        const ACTION_FQN_PP =
            "com.sap.gateway.srvd.zpp_ui_zuplsx.v0001.uploadFromExcel";

        const _str = (v) => (!v ? "" : `${v}`);
        const _num = (v) => (!v || v === "" ? "0" : `${v}`);
        const _date = (v) => {
            if (!v || v.length !== 10) return "";
            // DD/MM/YYYY -> YYYYMMDD
            return `${v.substring(6)}${v.substring(3, 5)}${v.substring(0, 2)}`;
        };

        return {
            /**
             * Kiểm tra file đã post chưa (entity UploadHistory) bằng ODataModel V4.
             * @param {sap.ui.model.odata.v4.ODataModel} oModel
             * @param {string} sFileName
             * @returns {Promise<void>} resolve nếu chưa tồn tại, reject nếu đã có
             */
            checkFileExists: function (oModel, sFileName) {
                return new Promise((resolve, reject) => {
                    const oListBinding = oModel.bindList("/UploadHistory", null, null, [
                        new Filter("filename", FilterOperator.EQ, sFileName),
                    ]);
                    oListBinding
                        .requestContexts(0, 1)
                        .then((aContexts) => {
                            if (aContexts && aContexts.length > 0) {
                                reject(
                                    new Error(
                                        `File ${sFileName} đã được post, vui lòng chọn File khác!`
                                    )
                                );
                            } else {
                                resolve();
                            }
                        })
                        .catch(() =>
                            reject(new Error("Lỗi kiểm tra lịch sử file. Vui lòng thử lại."))
                        );
                });
            },

            groupDataByDocId: function (aDataUpload) {
                const listDoc = new Map();
                aDataUpload.forEach((item) => {
                    if (!listDoc.has(item.id_doc)) listDoc.set(item.id_doc, []);
                    listDoc.get(item.id_doc).push(item);
                });
                return listDoc;
            },

            buildAllDocs: function (groupedDocs) {
                const aDocs = [];
                for (const [idDoc, items] of groupedDocs.entries()) {
                    const item_doc = items.map((item, index) => ({
                        idline: `${index + 1}`,
                        postingkey: _str(item.postingkey),
                        negativeposting: _str(item.negativeposting),
                        account: _str(item.account),
                        mainassetnumber: _str(item.mainassetnumber),
                        subassetnumber: _str(item.subassetnumber),
                        specialglaccount: _str(item.specialglaccount),
                        assettransactiontype: _str(item.assettransactiontype),
                        amountinlocalcurrency: _str(item.amountinlocalcurrency),
                        transactioncurrency: _str(item.currency),
                        amountindoumentcurrency: _num(item.amountindoumentcurrency),
                        taxbaseamount: _num(item.taxbaseamount),
                        localtaxbaseamount: _num(item.localtaxbaseamount),
                        exchangerate: _num(item.exchangerate),
                        assignment: _str(item.assignment),
                        costcenter: _str(item.costcenter),
                        profitcenter: _str(item.profitcenter),
                        internalorder: _str(item.internalorder),
                        wbselement: _str(item.wbselement),
                        businessarea: _str(item.businessarea),
                        assetvaluedate: _date(item.assetvaluedate),
                        itemtext: _str(item.itemtext),
                        longtext: _str(item.longtext),
                        overrideglaccount: _str(item.overrideglaccount),
                        taxcode: _str(item.taxcode),
                        segment: _str(item.segment),
                        paymentterms: _str(item.paymentterms),
                        paymentblockreason: _str(item.paymentblockreason),
                        paymentmethod: _str(item.paymentmethod),
                        baselinedate: _date(item.baselinedate),
                        valuedate: _date(item.valuedate),
                        netduedate: _date(item.netduedate),
                        contractnumber: _str(item.contractnumber),
                        contracttype: _str(item.contracttype),
                        housebank: _str(item.housebank),
                        bankaccountid: _str(item.bankaccountid),
                        invoicerefnum: _str(item.invoicerefnum),
                        invoicereffiscalyear: _str(item.invoicereffiscalyear),
                        invoicereflineitem: _str(item.invoicereflineitem),
                        purchasingno: _str(item.purchasingno),
                        purchasingitem: _str(item.purchasingitem),
                        saleorder: _str(item.saleorder),
                        saleorderitem: _str(item.saleorderitem),
                        customer: _str(item.customer),
                        cusgroup: _str(item.cusgroup),
                        alternativepayee: _str(item.alternativepayee),
                        division: _str(item.division),
                        distributionchannel: _str(item.distributionchannel),
                        salesorganization: _str(item.salesorganization),
                        salesoffice: _str(item.salesoffice),
                        salesemployee: _str(item.salesemployee),
                        salesgroup: _str(item.salesgroup),
                        materialgroup: _str(item.materialgroup),
                        product: _str(item.product),
                        unit: _str(item.unit),
                        baseunit: _str(item.baseunit),
                        quantity: _str(item.quantity),
                        name1: _str(item.name1),
                        name2: _str(item.name2),
                        name3: _str(item.name3),
                        name4: _str(item.name4),
                        mst: _str(item.mst),
                        city: _str(item.city),
                        country: _str(item.country),
                        vatregno: _str(item.vatregno),
                        ref1: _str(item.ref1),
                        ref2: _str(item.ref2),
                        ref3: _str(item.ref3),
                        namecus1: _str(item.namecus1),
                        namecus2: _str(item.namecus2),
                        namecus3: _str(item.namecus3),
                        namecus4: _str(item.namecus4),
                        mstcus: _str(item.mstcus),
                        citycus: _str(item.citycus),
                        countrycus: _str(item.countrycus),
                        countrygl: _str(item.countrygl),
                        plant: _str(item.plant),
                        productgroupmot: _str(item.productgroupmot),
                        producttype: _str(item.producttype),
                        orderid: _str(item.orderid),
                        material: _str(item.material),
                    }));

                    const header = items[0];
                    aDocs.push({
                        idDoc: idDoc,
                        companycode: header.companycode,
                        documentdate: _date(header.documentdate),
                        postingdate: _date(header.postingdate),
                        documenttype: header.documenttype,
                        currency: header.currency,
                        headertext: header.headertext,
                        referencedoc: header.referencedoc,
                        headerref1: header.headerref1,
                        toItem: item_doc,
                    });
                }
                return aDocs;
            },

            /**
             * Gọi OData V4 bound action uploadFromExcel — TẤT CẢ doc trong 1 call.
             * Cú pháp đã verify (Cách 1): bindContext("/UploadHistory/<FQN>(...)") + execute.
             * @returns {Promise<{results:Array, messages:Array}>}
             */
            callActionUpload: async function (oModel, sFileName, isUpdate, testMode, aDocs) {
                const sPayloadJson = JSON.stringify({
                    isupdate: isUpdate,
                    testmode: testMode,
                    filename: sFileName,
                    useremail: await this.getUserEmail(),
                    doc: aDocs,
                });

                const oOperation = oModel.bindContext("/UploadHistory/" + ACTION_FQN + "(...)");
                oOperation.setParameter("PayloadJson", sPayloadJson);
                await oOperation.execute();

                const oResult = oOperation.getBoundContext().getObject();
                const aResults = (oResult && oResult.value) || [];

                // Field PascalCase từ ZD_FIDocUploadResult
                const aMessages = aResults.map((value) => ({
                    type: value.Type,
                    title: value.Message,
                    group: `ID Doc ${value.IdDoc}`,
                }));

                return { results: aResults, messages: aMessages };
            },

            // ════════════════════════════════════════════════════════
            // PP — Upload Lệnh Sản Xuất (service zpp_ui_zuplsx_o4)
            // ════════════════════════════════════════════════════════
            buildPpRows: function (aItems) {
                return (aItems || []).map((item) => ({
                    clientRowId: _str(item.clientRowId),
                    idDoc: _str(item.id_doc),
                    orderType: _str(item.ordertype),
                    productionPlant: _str(item.productionplant),
                    totalQty: _str(item.totalqty),
                    baseUnit: _str(item.baseunit),
                    material: _str(item.material),
                    productionVersion: _str(item.productionversion),
                    dateStart: _str(item.datestart),
                    dateEnd: _str(item.dateend),
                    productionOrder: _str(item.productionorder),
                    saleOrder: _str(item.saleorder),
                    saleOrderItem: _str(item.saleorderitem),
                    longText: _str(item.longtext),
                }));
            },

            checkFileExistsPP: function (oModel, sFileName) {
                return new Promise((resolve, reject) => {
                    const oListBinding = oModel.bindList("/UploadHistory", null, null, [
                        new Filter("Filename", FilterOperator.EQ, sFileName),
                    ]);
                    oListBinding
                        .requestContexts(0, 1)
                        .then((aContexts) => {
                            if (aContexts && aContexts.length > 0) {
                                reject(new Error(`File ${sFileName} đã được post, vui lòng chọn File khác!`));
                            } else {
                                resolve();
                            }
                        })
                        .catch(() =>
                            reject(new Error("Lỗi kiểm tra lịch sử file. Vui lòng thử lại."))
                        );
                });
            },

            /**
             * Gọi OData V4 bound action uploadFromExcel của service PP.
             * @param {sap.ui.model.odata.v4.ODataModel} oModel model "pp"
             * @param {string} sFileName tên file (chỉ để log message, PP không lưu filename)
             * @param {string} sTestMode "X" = chỉ validate, "" = tạo lệnh thật
             * @param {Array} aRows từ buildPpRows
             * @returns {Promise<{results:Array, messages:Array}>}
             *          results item: { ClientRowId, IdDoc, Type, Message, ProductionOrder }
             */
            callActionUploadPP: async function (oModel, sFileName, sTestMode, aRows) {
                if (!oModel) {
                    throw new Error(
                        "Chưa cấu hình model 'pp' (service zpp_ui_zuplsx_o4) trong manifest.json"
                    );
                }

                const sPayloadJson = JSON.stringify({
                    filename: sFileName,
                    testmode: sTestMode,
                    useremail: await this.getUserEmail(),
                    rows: aRows,
                });

                const oOperation = oModel.bindContext(
                    "/UploadHistory/" + ACTION_FQN_PP + "(...)"
                );
                oOperation.setParameter("PayloadJson", sPayloadJson);
                await oOperation.execute();

                const oResult = oOperation.getBoundContext().getObject();
                const aResults = (oResult && oResult.value) || [];

                // Field PascalCase từ ZD_PPOrdUploadResult
                const aMessages = aResults.map((value) => ({
                    type: value.Type,
                    title: value.Message,
                    group: `ID ${value.IdDoc}`,
                }));

                return { results: aResults, messages: aMessages };
            },
            // ════════════════════════════════════════════════════════
            // GR — Upload Phiếu Nhập Kho (service zmm_ui_pogr_o4)
            // ════════════════════════════════════════════════════════

            buildGrDocs: function (aItems) {
                const mGr = new Map();
                aItems.forEach((row) => {
                    const sKey = row.gr_number || "";
                    if (!mGr.has(sKey)) {
                        mGr.set(sKey, {
                            grnumber: sKey,
                            documentdate: row.document_date || "",
                            movementtype: row.movement_type || "101",
                            items: [],
                        });
                    }
                    mGr.get(sKey).items.push({
                        ponumber: _str(row.po_number),
                        poitem: _str(row.po_item),
                        quantity: _str(row.receive_qty),
                        baseunit: _str(row.unit),
                        batch: _str(row.batch),
                        storagelocation: _str(row.storage_location),
                    });
                });
                return Array.from(mGr.values());
            },


            callActionUploadGR: async function (oModel, sFileName, bTestMode, aDocs) {
                const ACTION_FQN_GR =
                    "com.sap.gateway.srvd.zmm_ui_pogr_o4.v0001.uploadExcel";

                const sPayloadJson = JSON.stringify({
                    filename: sFileName,
                    useremail: await this.getUserEmail(),
                    doc: aDocs,
                });

                const oOperation = oModel.bindContext(
                    "/GrUpload/" + ACTION_FQN_GR + "(...)"
                );
                oOperation.setParameter("payload_json", sPayloadJson);
                oOperation.setParameter("mapping_id", "POGR001");
                oOperation.setParameter("testmode", bTestMode);
                await oOperation.execute();

                const oResult = oOperation.getBoundContext().getObject();
                return oResult || {};
            },

            /**
 * Đọc quyền upload của user hiện tại. Action đặt ở service GR nhưng trả quyền
 * của cả ba phân hệ, nên FI và PP dùng chung, không cần sửa service của chúng.
 */
            loadMyAuth: async function (oGrModel) {
                if (!oGrModel) return {};
                const oOperation = oGrModel.bindContext(
                    "/GrUpload/com.sap.gateway.srvd.zmm_ui_pogr_o4.v0001.getMyAuth(...)"
                );
                oOperation.setParameter("user_email", await this.getUserEmail());   // ← THÊM
                await oOperation.execute();
                return oOperation.getBoundContext().getObject() || {};
            },

            _sUserEmail: null,

            /**
             * Email người đang đăng nhập launchpad. Cache vì chỉ cần hỏi 1 lần.
             * Chạy ngoài launchpad (BAS preview) trả rỗng — backend sẽ từ chối kèm lý do.
             */
            getUserEmail: async function () {
                if (this._sUserEmail !== null) return this._sUserEmail;
                this._sUserEmail = "";
                try {
                    if (typeof sap !== "undefined" && sap.ushell && sap.ushell.Container) {
                        const oUserInfo = await sap.ushell.Container.getServiceAsync("UserInfo");
                        this._sUserEmail = (oUserInfo.getEmail() || oUserInfo.getId() || "").toLowerCase();
                    }
                } catch (e) {
                    this._sUserEmail = "";
                }
                return this._sUserEmail;
            },


        };
    }
);