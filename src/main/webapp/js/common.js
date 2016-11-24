
/**
 * 公共excel导出   针对easyui datagrid导出
 * @param _$datagrid  要导出的datagrid  是个jquery对象
 * @param _sheetName  导出的excel的sheet名称
 * @param _fileName    导出的excel文件名
 * @param _url  导出的后台地址
 * @param _params  传递额外参数对象 请确保key名称不会和查询参数重复
 */
$.exportExcel = function (_$datagrid, _sheetName, _fileName, _url, _params) {
    $.messager.confirm('确认', '您确认想要导出吗？', function (r) {
        if (r) {
            //获取 datagrid 配置需要传递的参数
            var queryParams = _$datagrid.datagrid("options").queryParams;
            //获取 datagrid 分页参数
            var pagination = _$datagrid.datagrid('getPager').data("pagination");
            var pager = pagination ? pagination.options : null;
            if (pager) {
                queryParams.page = pager.pageNumber;
                queryParams.rows = pager.pageSize;
                if (_url.endWith("1")) {
                    _fileName = _fileName + "第" + queryParams.page + "页";
                    _sheetName = _sheetName + "第" + queryParams.page + "页";
                }
            } else {//如果分页不存在  这使用 导出全部
                _url = _url.substring(0, _url.length - 1) + '0';
            }
            var data = JSON.stringify({
                TitleInfo: _$datagrid.datagrid("options").columns,
                SheetName: _sheetName,
                SaveName: _fileName
            });
            //添加 标题行数据
            queryParams.excelTitle = data;
            //额外参数 添加 进入查询参数中
            for (var key in _params) {
                if (queryParams[key] || queryParams[key] == 0) {

                } else {
                    queryParams[key] = _params[key];
                }
            }
            //构建form 表单
            var form = document.createElement("form");
            form.action = _url;
            form.method = "post";
            form.style.display = "none";
            form.target = "_blank";
            //创建 表单元素
            for (var key in queryParams) {
                var input = document.createElement("input");
                input.name = key;
                input.setAttribute("value", queryParams[key]);
                form.appendChild(input);
            }
            var optS = document.createElement("input");
            optS.type = "submit";
            optS.name = "postsubmit";
            form.appendChild(optS);
            document.body.appendChild(form);
            form.submit();
            document.body.removeChild(form);
        }
    });
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
