sap.ui.define(
  [
    "sap/m/MessageItem",
    "sap/m/MessageView",
    "sap/m/Dialog",
    "sap/m/Button",
    "sap/m/Bar",
    "sap/m/Title",
    "sap/m/MessageStrip",
    "sap/m/VBox",
    "sap/m/FlexItemData",
    "sap/m/MessageToast",
    "sap/ui/core/IconPool",
    "sap/ui/core/library",
    "sap/ui/model/json/JSONModel",
  ],
  function (
    MessageItem,
    MessageView,
    Dialog,
    Button,
    Bar,
    Title,
    MessageStrip,
    VBox,
    FlexItemData,
    MessageToast,
    IconPool,
    coreLibrary,
    JSONModel,
  ) {
    "use strict";

    const TitleLevel = coreLibrary.TitleLevel;

    let oMsgView = null;
    let oDialog = null;
    let oStrip = null;
    let oHeaderTitle = null;
    let oCopyBtn = null;
    let oCloseBtn = null;
    let sDocNumbers = "";

    /**
     * getText trả về CHÍNH KEY khi thiếu entry trong i18n.properties,
     * nên `getText(k) || fallback` không bao giờ chạy tới fallback
     * mà đẩy thẳng tên key ra màn hình. Hàm này so với key nên fallback
     * mới thật sự có tác dụng.
     */
    function _txt(oBundle, sKey, aArgs, sFallback) {
      if (!oBundle) return sFallback;
      const s = oBundle.getText(sKey, aArgs);
      return !s || s === sKey ? sFallback : s;
    }

    /**
     * Gom kết quả từ mảng results của action.
     * FI trả AccountingDocument, PP trả ProductionOrder -> lấy cái nào có.
     * Không truyền results thì rơi về đếm theo messages (chỉ có ok/err).
     */
    /**
     * Gom kết quả THEO CHỨNG TỪ, không theo số message:
     * một chứng từ có thể trả về nhiều message (BAPI trả nhiều dòng lỗi).
     * Chứng từ tính là lỗi nếu có BẤT KỲ message nào Type = Error.
     * Mọi Type khác (Success, Information, Warning) đều tính là OK -
     * chế độ Check trả Type = Information nên không được bỏ sót.
     */
    function _summarize(aMessages, aResults) {
      const bUseResults = !!(aResults && aResults.length);
      const aRes = bUseResults ? aResults : aMessages || [];
      const oDocs = {};
      const aOrder = [];

      aRes.forEach(function (o, i) {
        let sKey;
        if (bUseResults) {
          sKey =
            o.IdDoc !== undefined && o.IdDoc !== null
              ? String(o.IdDoc)
              : String(o.ClientRowId !== undefined ? o.ClientRowId : i);
        } else {
          sKey = String(o.group !== undefined ? o.group : i);
        }

        if (!oDocs[sKey]) {
          oDocs[sKey] = { err: false, doc: "" };
          aOrder.push(sKey);
        }

        if (String(o.Type || o.type || "") === "Error") {
          oDocs[sKey].err = true;
        }

        const sDoc = o.AccountingDocument || o.ProductionOrder;
        // bỏ giá trị toàn số 0 (chưa tạo được chứng từ)
        if (sDoc && String(sDoc).replace(/0/g, "") !== "") {
          oDocs[sKey].doc = String(sDoc);
        }
      });

      let iOk = 0;
      let iErr = 0;
      const aDocs = [];
      aOrder.forEach(function (sKey) {
        if (oDocs[sKey].err) iErr++;
        else iOk++;
        if (oDocs[sKey].doc) aDocs.push(oDocs[sKey].doc);
      });

      return {
        ok: iOk,
        err: iErr,
        total: iOk + iErr,
        docs: aDocs,
        docList: aDocs.join(", "),
      };
    }

    /** Quyết định nội dung, màu và tiêu đề của dải tổng hợp */
    function _buildStrip(oBundle, oSum, bTestMode) {
      const sDocPart = oSum.docList ? ": " + oSum.docList : "";

      if (bTestMode) {
        if (oSum.err > 0) {
          return {
            type: "Warning",
            state: "Warning",
            title: _txt(
              oBundle,
              "RESULT_TITLE_CHECK_ERR",
              null,
              "Kiểm tra: có lỗi",
            ),
            text: _txt(
              oBundle,
              "RESULT_STRIP_CHECK_ERR",
              [oSum.err, oSum.total],
              "CHẾ ĐỘ KIỂM TRA - " +
                oSum.err +
                "/" +
                oSum.total +
                " chứng từ có lỗi. Chưa có gì được ghi vào hệ thống.",
            ),
          };
        }
        return {
          type: "Information",
          state: "Information",
          title: _txt(
            oBundle,
            "RESULT_TITLE_CHECK_OK",
            null,
            "Kiểm tra: hợp lệ",
          ),
          text: _txt(
            oBundle,
            "RESULT_STRIP_CHECK_OK",
            [oSum.ok, oSum.total],
            "CHẾ ĐỘ KIỂM TRA - " +
              oSum.ok +
              "/" +
              oSum.total +
              " chứng từ hợp lệ. Chưa có gì được ghi vào hệ thống, bấm Post để hạch toán thật.",
          ),
        };
      }

      if (oSum.err === 0 && oSum.ok > 0) {
        return {
          type: "Success",
          state: "Success",
          title: _txt(oBundle, "RESULT_TITLE_POST_OK", null, "Post thành công"),
          text: _txt(
            oBundle,
            "RESULT_STRIP_POST_OK",
            [oSum.ok, oSum.total, oSum.docList],
            "POST THÀNH CÔNG - đã tạo " +
              oSum.ok +
              "/" +
              oSum.total +
              " chứng từ" +
              sDocPart +
              ".",
          ),
        };
      }

      if (oSum.ok > 0 && oSum.err > 0) {
        return {
          type: "Warning",
          state: "Warning",
          title: _txt(
            oBundle,
            "RESULT_TITLE_POST_PARTIAL",
            null,
            "Post một phần",
          ),
          text: _txt(
            oBundle,
            "RESULT_STRIP_POST_PARTIAL",
            [oSum.ok, oSum.total, oSum.err, oSum.docList],
            "POST MỘT PHẦN - " +
              oSum.ok +
              "/" +
              oSum.total +
              " chứng từ được tạo" +
              (oSum.docList ? " (" + oSum.docList + ")" : "") +
              ", " +
              oSum.err +
              " chứng từ lỗi. Xem chi tiết bên dưới.",
          ),
        };
      }

      return {
        type: "Error",
        state: "Error",
        title: _txt(oBundle, "RESULT_TITLE_POST_FAIL", null, "Post thất bại"),
        text: _txt(
          oBundle,
          "RESULT_STRIP_POST_FAIL",
          [oSum.err, oSum.total],
          "POST THẤT BẠI - không chứng từ nào được tạo (" +
            oSum.err +
            "/" +
            oSum.total +
            " lỗi). Sửa file rồi post lại.",
        ),
      };
    }

    function _copyToClipboard(sText) {
      if (navigator.clipboard && navigator.clipboard.writeText) {
        return navigator.clipboard.writeText(sText);
      }
      const oTa = document.createElement("textarea");
      oTa.value = sText;
      document.body.appendChild(oTa);
      oTa.select();
      const bOk = document.execCommand("copy");
      document.body.removeChild(oTa);
      return bOk ? Promise.resolve() : Promise.reject();
    }

    function _createDialog(oBundle) {
      const oBackButton = new Button({
        icon: IconPool.getIconURI("nav-back"),
        visible: false,
        press: () => {
          oMsgView.navigateBack();
          oBackButton.setVisible(false);
        },
      });

      oMsgView = new MessageView({
        showDetailsPageHeader: false,
        itemSelect: () => oBackButton.setVisible(true),
        items: {
          path: "/",
          template: new MessageItem({
            type: "{type}",
            title: "{title}",
            groupName: "{group}",
          }),
        },
        groupItems: true,
        layoutData: new FlexItemData({ growFactor: 1, baseSize: "0%" }),
      });

      // MỚI: dải tổng hợp, luôn nằm trên danh sách message
      oStrip = new MessageStrip({
        showIcon: true,
        showCloseButton: false,
        class: "sapUiTinyMargin",
      });

      oHeaderTitle = new Title({
        text: _txt(oBundle, "RESULTS_TITLE", null, "Kết quả"),
        level: TitleLevel.H1,
      });

      // MỚI: copy nhanh số chứng từ để dán sang FB03 / CO03
      oCopyBtn = new Button({
        icon: "sap-icon://copy",
        visible: false,
        press: function () {
          if (!sDocNumbers) {
            MessageToast.show(
              _txt(oBundle, "COPY_NO_VALUE", null, "Không có giá trị để copy"),
            );
            return;
          }
          _copyToClipboard(sDocNumbers)
            .then(function () {
              MessageToast.show(
                _txt(
                  oBundle,
                  "COPY_SUCCESS",
                  [sDocNumbers],
                  "Đã copy: " + sDocNumbers,
                ),
              );
            })
            .catch(function () {
              MessageToast.show(
                _txt(oBundle, "COPY_FAIL", null, "Copy thất bại"),
              );
            });
        },
      });

      oCloseBtn = new Button({
        text: _txt(oBundle, "CLOSE", null, "Đóng"),
        press: () => oDialog.close(),
      });

      oDialog = new Dialog({
        resizable: true,
        // VBox PHẢI có height + MessageView phải grow, nếu không
        // MessageView cao 0px và danh sách message biến mất.
        content: new VBox({
          height: "100%",
          items: [oStrip, oMsgView],
        }),
        state: "Information",
        beginButton: oCopyBtn,
        endButton: oCloseBtn,
        customHeader: new Bar({
          contentLeft: [oBackButton],
          contentMiddle: [oHeaderTitle],
        }),
        contentHeight: "55%",
        contentWidth: "55%",
        verticalScrolling: false,
      });
    }

    return {
      /**
       * Hiện kết quả Check/Post.
       * @param {object} oController controller gọi
       * @param {Array}  allMessages [{type,title,group}] đổ vào MessageView
       * @param {string} sTestMode   "X" = Check, "" = Post thật
       * @param {string} sPostBtnId  id nút Post để bật/tắt
       * @param {Array}  [aResults]  results gốc của action (có AccountingDocument /
       *                             ProductionOrder). Truyền vào thì dải tổng hợp
       *                             mới liệt kê được số chứng từ.
       * @returns {object} {ok, err, total, docs, docList, state}
       */
      show: function (
        oController,
        allMessages,
        sTestMode,
        sPostBtnId,
        aResults,
      ) {
        let oBundle = null;
        try {
          oBundle = oController.getView().getModel("i18n").getResourceBundle();
        } catch (e) {
          /* không có i18n thì dùng fallback tiếng Việt trong code */
        }

        if (!oDialog) _createDialog(oBundle);

        const bTest = sTestMode === "X";
        const oSum = _summarize(allMessages, aResults);
        const oCfg = _buildStrip(oBundle, oSum, bTest);

        oStrip.setType(oCfg.type);
        oStrip.setText(oCfg.text);
        oHeaderTitle.setText(oCfg.title);
        oDialog.setState(oCfg.state);

        oCopyBtn.setText(
          _txt(oBundle, "COPY_DOC_NUMBERS", null, "Copy số chứng từ"),
        );

        sDocNumbers = oSum.docList;
        oCopyBtn.setVisible(!bTest && oSum.docs.length > 0);

        const oMsgModel = new JSONModel();
        oMsgModel.setData(allMessages);
        oMsgModel.setSizeLimit(allMessages.length || 1);
        oMsgView.setModel(oMsgModel);
        oMsgView.navigateBack();
        oDialog.open();

        const oPostButton = oController.byId(sPostBtnId);
        if (oPostButton) {
          if (oSum.err > 0) {
            oPostButton.setEnabled(false);
          } else if (!bTest) {
            oPostButton.setEnabled(false);
          } else {
            oPostButton.setEnabled(true);
          }
        }

        oSum.state = oCfg.state;
        return oSum;
      },
    };
  },
);
