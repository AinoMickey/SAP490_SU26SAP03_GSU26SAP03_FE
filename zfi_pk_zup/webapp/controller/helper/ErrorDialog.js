sap.ui.define([
    "sap/ui/model/json/JSONModel",
    "sap/m/MessageItem",
    "sap/m/MessageView",
    "sap/m/Button",
    "sap/m/Dialog",
    "sap/m/Bar",
    "sap/m/Title",
    "sap/ui/core/IconPool",
    "sap/ui/core/library",
], function (JSONModel, MessageItem, MessageView, Button, Dialog, Bar, Title, IconPool, coreLibrary) {
    "use strict";
    const TitleLevel = coreLibrary.TitleLevel;
    return {
        handleErrorDialog: function (pr_aMockMessages, that) {
            let oMessageTemplate = new MessageItem({
                type: "{type}",
                title: "{title}",
                description: "{description}",
                subtitle: "{subtitle}",
            });
            let oModel = new JSONModel();
            oModel.setData(pr_aMockMessages);
            that.oMessageView = new MessageView({
                showDetailsPageHeader: false,
                itemSelect: function () {
                    oBackButton.setVisible(true);
                },
                items: {
                    path: "/",
                    template: oMessageTemplate,
                },
            });
            that.oMessageView.setModel(oModel);
            let oBackButton = new Button({
                icon: IconPool.getIconURI("nav-back"),
                visible: false,
                press: function () {
                    that.oMessageView.navigateBack();
                    this.setVisible(false);
                },
            });
            that.oDialog_Error = new Dialog({
                resizable: true,
                content: that.oMessageView,
                state: "Error",
                beginButton: new Button({
                    press: function () {
                        this.getParent().close();
                    },
                    text: that.getView().getModel("i18n").getResourceBundle().getText("CLOSE") || "Đóng",
                }),
                customHeader: new Bar({
                    contentLeft: [oBackButton],
                    contentMiddle: [
                        new Title({
                            text: that.getView().getModel("i18n").getResourceBundle().getText("ERROR_DIALOG_TITLE") || "Thông báo lỗi",
                            level: TitleLevel.H1,
                        }),
                    ],
                }),
                contentHeight: "50%",
                contentWidth: "50%",
                verticalScrolling: false,
            });
            that.oMessageView.navigateBack();
            that.oDialog_Error.open();
        },
    };
});