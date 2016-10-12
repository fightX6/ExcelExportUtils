
/**
 * 公共excel导出   针对easyui datagrid导出
 * @param _$datagrid  要导出的datagrid  是个jquery对象
 * @param _sheetName  导出的excel的sheet名称
 * @param _fileName    导出的excel文件名
 * @param _url  导出的后台地址
 */
$.exportExcel = function(_$datagrid,_sheetName,_fileName,_url){
    var data = JSON.stringify({
        TitleInfo: _$datagrid.datagrid("options").columns,
        SheetName: _sheetName,
        SaveName: _fileName
    });
    var form = document.createElement("form");
    form.action = _url;
    form.method = "post";
    form.style.display="none";
    form.target = "_blank";
    var opt = document.createElement("input");
    opt.name = "excelTitle";
    opt.setAttribute("value",data);
    form.appendChild(opt);
    var optS = document.createElement("input");
    optS.type = "submit";
    optS.name = "postsubmit";
    form.appendChild(optS);
    document.body.appendChild(form);
    form.submit();
    document.removeChild(form);
};
/**
 * 字符串添加startWith方法
 * @param str
 * @returns {boolean}
 */
String.prototype.startWith = function (str) {
    var reg = new RegExp("^" + str);
    return reg.test(this);
}
/**
 * 字符串添加endWith方法
 * @param str
 * @returns {boolean}
 */
String.prototype.endWith = function (str) {
    var reg = new RegExp(str + "$");
    return reg.test(this);
}
