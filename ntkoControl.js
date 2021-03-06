"use strict";

function NtkoControl() {
    this.ctx = window.pageContext; // 上下文
    this.browser = {name: "", version: "0"}; //  浏览器 ie/chrome/firefox
    this.isIE = false;
    this.tangerOcx = null; // html object : document.getElementById("TANGER_OCX")

    this.fileName = "";
    this.isOpenURLReadOnly = false;

    this.formData = undefined; // 表单值，用于套红

    this._onDocumentOpened = function (str, obj) {
        var thiz = window.ntko;
        if (str) {
            var pos = str.lastIndexOf("/");
            if (pos < 0) pos = str.lastIndexOf("\\");
            thiz.fileName = (pos >= 0) ? str.substr(pos + 1) : str;
        } else {
            thiz.fileName = "";
        }

        if (thiz.isOpenURLReadOnly && thiz.tangerOcx.IsOpenFromUrl) {
            thiz.protect(thiz.protectKey);
        } else {
            thiz.tangerOcx.IsShowToolMenu = !thiz.isProtect(); // 重置tool menu
        }

        thiz.trackRevisions(true); // 尝试进入痕迹模式
        thiz.showRevisions(false); // 隐藏痕迹

        if (thiz.onDocumentOpened) thiz.onDocumentOpened(thiz);
    };

    this._customToolBarCmd = function (btnIdx) {
        var thiz = window.ntko;
        thiz.tangerOcx.toolbars = !thiz.tangerOcx.toolbars;
    };

    this._onSaveToUrl = function (type, code, html) {
        try {
            var data = html;
            if (typeof data !== "object") data = JSON.parse(data);
            if (typeof data !== "object") data = {code: -1, msg: "保存文件出错了！"};
        } catch (e) {
            data = {code: -1, msg: "保存文件出错了！"};
        }

        var thiz = window.ntko;
        if (thiz.onSaveToUrl) thiz.onSaveToUrl(data);
    };

    this._AddTemplateFromURL = function (type, code, html) {
        var thiz = window.ntko;
        var doc = thiz.tangerOcx.ActiveDocument;

        var BookMarkName = "正文";
        if (!doc.BookMarks.Exists(BookMarkName)) {
            var sel = doc.Application.Selection;
            sel.WholeStory();
            sel.Paste();
            alert('Word 模板中不存在名称为："' + BookMarkName + '"的书签！');
            return;
        }

        doc.BookMarks(BookMarkName).Range.Paste();

        if (thiz.formData) {
            for (var k in thiz.formData) {
                if (doc.BookMarks.Exists(k)) thiz.tangerOcx.SetBookmarkValue(k, thiz.formData[k]);
            }
        }

        thiz.acceptRevisions(true);
    };

    //author: meizz
    this.formatDate = function (date, fmt) {
        var o = {
            "M+": date.getMonth() + 1, //月份
            "d+": date.getDate(), //日
            "h+": date.getHours(), //小时
            "m+": date.getMinutes(), //分
            "s+": date.getSeconds(), //秒
            "q+": Math.floor((date.getMonth() + 3) / 3), //季度
            "S": date.getMilliseconds() //毫秒
        };
        if (/(y+)/.test(fmt)) fmt = fmt.replace(RegExp.$1, (date.getFullYear() + "").substr(4 - RegExp.$1.length)); // 年
        for (var k in o)
            if (new RegExp("(" + k + ")").test(fmt)) fmt = fmt.replace(RegExp.$1, (RegExp.$1.length === 1) ? (o[k]) : (("00" + o[k]).substr(("" + o[k]).length)));
        return fmt;
    };

    this.detectBrowser = function () {
        var ua = navigator.userAgent.toLowerCase();
        var match;

        match = ua.match(/(msie\s|trident.*rv:)([\w.]+)/);
        if (match) return {name: "ie", version: match[2] || "0"};

        match = ua.match(/(chrome)\/([\w.]+)/);
        if (match) return {name: match[1] || "", version: match[2] || "0"};

        match = ua.match(/(firefox)\/([\w.]+)/);
        if (match) return {name: match[1] || "", version: match[2] || "0"};

        match = ua.match(/(opera).+version\/([\w.]+)/);
        if (match) return {name: match[1] || "", version: match[2] || "0"};

        match = ua.match(/version\/([\w.]+).*(safari)/);
        if (match) return {name: match[2] || "", version: match[1] || "0"};
        return {name: "", version: "0"};
    };

    // 显示 ntko 控件
    this.showTangerOcx = function () {
        var productKey = "product.key";
        var productCaption = "caption";
        var caption = productCaption + "(双击全屏)";

        var clsid;
        var codebase;

        if (this.browser.name === "ie") {
            if (window.navigator.platform === "Win32") {
                clsid = "A64E3073-2016-4baf-A89D-FFE1FAA10EC0";
                codebase = this.ctx + "/static/ntko/OfficeControl.cab#version=5.0.2.9";
            } else {
                clsid = "A64E3073-2016-4baf-A89D-FFE1FAA10EE0";
                codebase = this.ctx + "/static/ntko/OfficeControlX64.cab#version=5.0.2.9";
            }
        } else {
            clsid = "A64E3073-2016-4baf-A89D-FFE1FAA10EC0";
            codebase = this.ctx + "/static/ntko/OfficeControl.cab#version=5,0,2,9";
        }

        if (this.browser.name === "ie") {
            document.write('<object id="TANGER_OCX"'
                + ' classid="clsid:' + clsid + '"' + ' codebase="' + codebase + '"'
                + ' width="100%" height="100%">'
                + ' <param name="IsUseUTF8URL" value="-1">'
                + ' <param name="IsUseUTF8Data" value="-1">'
                + ' <param name="BorderStyle" value="1">'
                + ' <param name="BorderColor" value="14402205">'
                + ' <param name="TitlebarColor" value="15658734">'
                + ' <param name="TitlebarTextColor" value="0">'
                + ' <param name="Caption" value="' + caption + '">'
                + ' <param name="MakerCaption" value="">'
                + ' <param name="MakerKey" value="">'
                + ' <param name="ProductCaption" value="' + productCaption + '">'
                + ' <param name="ProductKey" value="' + productKey + '">'
                + ' </object>'
                + ' <script language="JScript" for="TANGER_OCX" event="OnDocumentOpened(str,obj)">'
                + '     window.ntko._onDocumentOpened(str,obj);'
                + ' </script>'
                + ' <script language="JScript" for="TANGER_OCX" event="OnCustomToolBarCommand(btnIdx)">'
                + '     window.ntko._customToolBarCmd(btnIdx);'
                + ' </script>');
        } else if (this.browser.name === "chrome") {
            document.write('<object id="TANGER_OCX"'
                + ' clsid="{' + clsid + '}"' + ' type="application/ntko-plug" codebase="' + codebase + '"'
                + ' width="100%" height="100%"'
                + ' ForOnDocumentOpened="ntko_control_onDocumentOpened"'
                + ' ForOnCustomToolBarCommand="ntko_control_customToolBarCmd"'
                + ' ForOnSaveToURL="ntko_control_onSaveToUrl"'
                + ' ForOnAddTemplateFromURL="ntko_control_AddTemplateFromURL"'
                + ' _IsUseUTF8URL="-1"'
                + ' _IsUseUTF8Data="-1"'
                + ' _BorderStyle="1"'
                + ' _BorderColor="14402205"'
                + ' _TitlebarColor="15658734"'
                + ' _TitlebarTextColor="0"'
                + ' _Caption="' + caption + '"'
                + ' _ProductKey="' + productKey + '"'
                + ' _ProductCaption="' + productCaption + '">'
                + ' </object>');
        } else if (this.browser.name === "firefox") {
            document.write('<object id="TANGER_OCX"'
                + ' clsid="{' + clsid + '}"' + ' type="application/ntko-plug" codebase="' + codebase + '"'
                + ' width="100%" height="100%"'
                + ' ForOnDocumentOpened="ntko_control_onDocumentOpened"'
                + ' ForOnCustomToolBarCommand="ntko_control_customToolBarCmd"'
                + ' ForOnSaveToURL="ntko_control_onSaveToUrl"'
                + ' ForOnAddTemplateFromURL="ntko_control_AddTemplateFromURL"'
                + ' _IsUseUTF8URL="-1"'
                + ' _IsUseUTF8Data="-1"'
                + ' _BorderStyle="1"'
                + ' _BorderColor="14402205"'
                + ' _TitlebarColor="15658734"'
                + ' _TitlebarTextColor="0"'
                + ' _Caption="' + caption + '"'
                + ' _ProductKey="' + productKey + '"'
                + ' _ProductCaption="' + productCaption + '">'
                + ' </object>');
        } else {
            alert("sorry, ntko/web印章暂不支持当前浏览器!");
        }
    };

    this.browser = this.detectBrowser();
    this.isIE = this.browser.name === "ie";
    this.protectKey = "123456l91";

    window.ntko_control_onDocumentOpened = this._onDocumentOpened; // 打开文档回调
    window.ntko_control_customToolBarCmd = this._customToolBarCmd; // 工具条按钮onclick回调
    window.ntko_control_onSaveToUrl = this._onSaveToUrl; // saveToUrl回调
    window.ntko_control_AddTemplateFromURL = this._AddTemplateFromURL; // 工具条按钮onclick回调
}

NtkoControl.prototype = {
    init: function (username) {
        this.tangerOcx = document.getElementById("TANGER_OCX");
        if (username) this.tangerOcx.WebUserName = username;

        if (window.navigator.platform === "Win32") {
            this.tangerOcx.AddDocTypePlugin(".pdf", "PDF.NtkoDocument", "4,0,0,6", this.ctx + "/static/ntko/ntkooledocall.cab", 51, true);
            this.tangerOcx.AddDocTypePlugin(".tif", "TIF.NtkoDocument", "4.0.0.6", this.ctx + "/static/ntko/ntkooledocall.cab", 52);
            this.tangerOcx.AddDocTypePlugin(".tiff", "TIF.NtkoDocument", "4.0.0.6", this.ctx + "/static/ntko/ntkooledocall.cab", 52);
        } else {
            this.tangerOcx.AddDocTypePlugin(".pdf", "PDF.NtkoDocument", "4,0,0,6", this.ctx + "/static/ntko/ntkooledocallX64.cab", 51, true);
            this.tangerOcx.AddDocTypePlugin(".tif", "TIF.NtkoDocument", "4.0.0.6", this.ctx + "/static/ntko/ntkooledocallX64.cab", 52);
            this.tangerOcx.AddDocTypePlugin(".tiff", "TIF.NtkoDocument", "4.0.0.6", this.ctx + "/static/ntko/ntkooledocallX64.cab", 52);
        }

        this.isOpenURLReadOnly = false;
        return this;
    },

    createNew: function (progId) {
        try {
            this.tangerOcx.CreateNew(progId);
        } catch (e) {
        }
    },

    createNewDoc: function () {
        this.createNew('Word.Document');
    },

    openFromLocal: function () {
        this.tangerOcx.ShowDialog(1);
    },

    openFromUrl: function (url, readonly) {
        if (this.tangerOcx == null) this.init();
        this.isOpenURLReadOnly = !!eval(readonly);

        if (this.tangerOcx.Caption != null) {
            this.tangerOcx.toolbars = false;

            if (this.tangerOcx.CustomToolBar !== true) { // 如果原来已经添加了，就不必要再次添加，否则出现重复的tool button
                this.tangerOcx.CustomToolBar = true;
                this.tangerOcx.AddCustomToolButton("工具栏", -1);
            }

            this.tangerOcx.Menubar = !this.isOpenURLReadOnly;

            try {
                if (url) {
                    this.tangerOcx.BeginOpenFromURL(this.ctx + "/" + url, false, this.isOpenURLReadOnly);
                } else {
                    this.createNew('Word.Document');
                }
            } catch (e) {
            }
        }
    },

    saveToLocal: function (fileName) {
        try {
            fileName = (typeof fileName === "string") ? fileName : fileName.getAttribute("SaveAsName"); // 可传递对象过来

            var dg = this.tangerOcx.ActiveDocument.Application.FileDialog(2);
            this.tangerOcx.Activate(true);
            if (fileName) fileName = (fileName.replace(/[/\\:*?"<>|]/g, "_")).replace(/^\s+|\s+$/g, "");
            dg.InitialFileName = fileName ? fileName : this.formatDate(new Date(), 'yyyyMMdd_hhmmss');
            if (dg.show() === -1) {
                this.tangerOcx.SaveToLocal(dg.selectedItems(1), true, false); //dg.Execute();
            }
        } catch (a) {
            this.tangerOcx.ShowDialog(3);
        }
    },

    saveToUrl: function (url) {
        this.trackRevisions(false);
        if (this.isOpenURLReadOnly) return; // 只读文件不保存到服务器

        var result;
        try {
            var params1 = {fileUrl: url ? url : "", docType: this.docType, rnd: Math.random()};
            var s = "";
            for (var key in params1) {
                s = s + "&" + encodeURIComponent(key) + "=" + encodeURIComponent(params1[key]);
            }
            params1 = s.replace(/%20/g, "+").substring(1); // 序列化后的字符串
            result = this.tangerOcx.saveToURL(this.ctx + "/api/ntko/save", "uploadFile", params1);

            // ie 是同步获取的，result 就是结果，
            // chrome / firefox 是异步的，得到结果后会自动调用 _onSaveToUrl
            if (this.isIE) {
                this._onSaveToUrl(null, null, result);
            }
        } catch (e) {
            result = {code: -1, msg: "保存文件出错了！"};
            this._onSaveToUrl(null, null, result);
        }
    },

    makeRed: function (templateUrl, data) {
        if (this.docType() === 1 || this.docType() === 6) {
            this.acceptRevisions(true);

            var curSel = this.tangerOcx.ActiveDocument.Application.Selection;
            curSel.WholeStory();
            curSel.Cut();

            this.formData = data;

            this.tangerOcx.AddTemplateFromURL(this.ctx + "/" + templateUrl);
            if (this.isIE) {
                this._AddTemplateFromURL();
            }
        } else {
            alert("不支持的文档内型");
        }
    },

    /*
        0   =   没有文档
        1   =   word
        2   =   Excel.Sheet或者 Excel.Chart
        3   =   PowerPoint.Show
        4   =   Visio.Drawing
        5   =   MSProject.Project
        6   =   WPS Doc
        7   =   Kingsoft.Sheet
        51  =   pdf
        100 =   其他文档类型
    */
    docType: function () {
        return this.tangerOcx.DocType;
    },

    isDocOpened: function () {
        return this.tangerOcx.ActiveDocument != null;
    },

    isProtect: function () {
        switch (this.tangerOcx.ActiveDocument.ProtectionType) {
            case -1:
                return false;

            case 0:
            case 1:
            case 2:
                return true;

            default:
                return true;
        }
    },

    protect: function (key) {
        try {
            if ((this.docType() === 1 || this.docType() === 6) && !this.isProtect()) {
                this.tangerOcx.ActiveDocument.Protect(3, false, key);
            }
            this.tangerOcx.IsShowToolMenu = !this.isProtect();
        } catch (e) {
        }
    },

    unProtect: function (key) {
        try {
            if ((this.docType() === 1 || this.docType() === 6) && this.isProtect()) {
                this.tangerOcx.ActiveDocument.UnProtect(key);
            }
            this.tangerOcx.IsShowToolMenu = !this.isProtect();
        } catch (e) {
        }
    },

    cut: function () {
        this.tangerOcx.ActiveDocument.application.Selection.Cut();
    },

    copy: function () {
        this.tangerOcx.ActiveDocument.application.Selection.Copy();
    },

    paste: function () {
        this.tangerOcx.ActiveDocument.application.Selection.PasteAndFormat(16);
    },

    print: function () {
        this.tangerOcx.PrintOut(true);
    },

    printPreview: function () {
        this.tangerOcx.PrintPreview();
    },

    exitPrintPreview: function () {
        this.tangerOcx.ExitPrintPreview();
    },

    // 启用/关闭痕迹修订
    trackRevisions: function (bool) {
        if (this.docType() === 1 || this.docType() === 6) {
            if (!this.isProtect()) this.tangerOcx.ActiveDocument.TrackRevisions = bool;
        }
    },

    // 显示或隐藏痕迹
    showRevisions: function (bool) {
        if (this.docType() === 1 || this.docType() === 6) {
            if (!this.isProtect()) this.tangerOcx.ActiveDocument.ShowRevisions = bool;
        }
    },

    // 接受或拒绝痕迹修订
    acceptRevisions: function (bool) {
        if (this.docType() === 1 || this.docType() === 6) {
            if (!this.isProtect()) {
                if (bool) {
                    this.tangerOcx.ActiveDocument.AcceptAllRevisions();
                } else if (this.docType() === 1) {
                    this.tangerOcx.ActiveDocument.Application.WordBasic.RejectAllChangesInDoc();
                } else { // docType === 6
                    this.tangerOcx.ActiveDocument.Revisions.RejectAll();
                }

                this.trackRevisions(false); // 关闭痕迹修订模式
            } else {
                alert("文档已被保护，接受或拒绝修订失败！"); // 前面 unprotect 失败了
            }
        }
    }

};

window.ntko = new NtkoControl();

