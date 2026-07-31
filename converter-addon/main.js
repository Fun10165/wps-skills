/*
 * 格式转换器 main.js
 * 批量将旧格式(.doc/.xls/.ppt)转换为新格式(.docx/.xlsx/.pptx)
 * 机制: 官方 WPS JS API (open.wps.cn) — 无轮询，加载即执行 + 工具栏按钮触发
 *
 * 格式代码 (VBA 标准, WPS 兼容):
 *   Word:  wdFormatXMLDocument = 12        -> .docx
 *   Excel: xlOpenXMLWorkbook   = 51        -> .xlsx
 *   PPT:   ppSaveAsOpenXMLPresentation = 24 -> .pptx
 *
 * 沙盒注意: 文件必须在 ~/Downloads 或 WPS 有完全磁盘访问的位置
 */

var CONV_LOG = [];

function convLog(msg) {
    CONV_LOG.push(String(msg));
    try { console.log('[格式转换器] ' + msg); } catch (e) {}
}

/* ---------------- 按输出扩展名推断应用 ---------------- */
function appForOutput(path) {
    var p = String(path).toLowerCase();
    if (p.indexOf('.docx') > -1 || p.indexOf('.doc') > -1) return 'wps';
    if (p.indexOf('.xlsx') > -1 || p.indexOf('.xls') > -1) return 'et';
    if (p.indexOf('.pptx') > -1 || p.indexOf('.ppt') > -1) return 'wpp';
    return '';
}

function fmtCodeForOutput(path) {
    var p = String(path).toLowerCase();
    if (p.indexOf('.docx') > -1) return 12;  // Word .docx
    if (p.indexOf('.xlsx') > -1) return 51;  // Excel .xlsx
    if (p.indexOf('.pptx') > -1) return 24;  // PPT .pptx
    return -1;
}

/* ---------------- 检测当前宿主应用 ---------------- */
function hostApp() {
    try { if (typeof wps !== 'undefined' && wps.Documents) return 'wps'; } catch (e) {}
    try { if (typeof et !== 'undefined' && et.Workbooks) return 'et'; } catch (e) {}
    try { if (typeof wpp !== 'undefined' && wpp.Presentations) return 'wpp'; } catch (e) {}
    return '';
}

/* ---------------- 单文件转换 ---------------- */
function convertOne(task, host) {
    var input = task.input, output = task.output;
    var overwrite = !!task.overwrite;

    try {
        if (host === 'wps') {
            wps.DisplayAlerts = 0;
            var doc = wps.Documents.Open(input, false);   // 第二个参数关闭“转换确认”对话框
            var code = fmtCodeForOutput(output);
            if (code === 12) {
                doc.SaveAs2(output, 12);                  // 新格式另存
            } else {
                doc.SaveAs2(output, code);
            }
            doc.Close(0);                                 // wdDoNotSaveChanges（已另存）
            return true;
        }

        if (host === 'et') {
            et.DisplayAlerts = false;
            var wb = et.Workbooks.Open(input);
            var code2 = fmtCodeForOutput(output);
            if (code2 > 0) {
                wb.SaveAs(output, code2);
            } else {
                wb.SaveAs(output);
            }
            wb.Close(false);
            return true;
        }

        if (host === 'wpp') {
            wpp.DisplayAlerts = 0;
            var pres = wpp.Presentations.Open(input);
            var code3 = fmtCodeForOutput(output);
            if (code3 > 0) {
                pres.SaveAs(output, code3);
            } else {
                pres.SaveAs(output);
            }
            pres.Close();
            return true;
        }
    } catch (err) {
        convLog('失败: ' + input + ' -> ' + (err && err.message ? err.message : err));
        return false;
    }
    convLog('跳过: ' + input + ' (不支持的输出格式)');
    return false;
}

/* ---------------- 批量执行 ---------------- */
function runConversion() {
    var host = hostApp();
    if (!host) {
        alert('格式转换器: 无法识别当前 WPS 组件。请打开 文字/表格/演示 后再试。');
        return;
    }

    var tasks = (window.CONV_TASKS || []).filter(function (t) {
        return appForOutput(t.output) === host;
    });

    if (tasks.length === 0) {
        alert('格式转换器: 当前组件(' + host + ')没有匹配的转换任务。请编辑 tasks.js 并重启 WPS。');
        return;
    }

    convLog('开始转换 ' + tasks.length + ' 个文件 (host=' + host + ')');
    var ok = 0, fail = 0;
    for (var i = 0; i < tasks.length; i++) {
        if (convertOne(tasks[i], host)) {
            ok++;
            convLog('完成: ' + tasks[i].input + ' -> ' + tasks[i].output);
        } else {
            fail++;
        }
    }
    convLog('结束: 成功 ' + ok + ' / 失败 ' + fail);
    alert('格式转换器: 完成 ' + ok + ' 个，失败 ' + fail + ' 个。');
}

/* ---------------- 入口 ---------------- */
function OnAddinLoad(ribbonUI) {
    convLog('加载中');
    // 延迟执行，等待 WPS JS API 就绪；最多重试 10 次
    var attempts = 0;
    var timer = setInterval(function () {
        attempts++;
        var ready = false;
        try {
            if (typeof wps !== 'undefined' && wps.Documents) ready = true;
            else if (typeof et !== 'undefined' && et.Workbooks) ready = true;
            else if (typeof wpp !== 'undefined' && wpp.Presentations) ready = true;
        } catch (e) {}
        if (ready) {
            clearInterval(timer);
            setTimeout(runConversion, 300);
        } else if (attempts > 10) {
            clearInterval(timer);
            convLog('API 未就绪，放弃自动执行（可用工具栏按钮手动触发）');
        }
    }, 500);
    return true;
}

function OnRunClick() {
    runConversion();
}
