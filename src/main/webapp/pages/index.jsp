<%@ page contentType="text/html;charset=UTF-8" language="java" %>
<html>
<head>
    <meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no">
    <link rel="stylesheet" type="text/css" href="${pageContext.request.contextPath}/js/easyui/themes/metro-default/easyui.css"/>
    <link rel="stylesheet" type="text/css" href="${pageContext.request.contextPath}/js/easyui/themes/icon.css"/>
    <link rel="stylesheet" type="text/css" href="${pageContext.request.contextPath}/js/easyui/themes/color.css"/>
    <script type="application/javascript">
        var BASE_URL = "${pageContext.request.contextPath}";
        /**
         * 构建url地址
         * @param uri
         * @returns {*}
         */
        function urlBuild(uri) {
            if (!uri) {
                throw "构建rest的url地址时，uri为空";
            }

            if (uri.substring(0, 1) === "/") {
                return BASE_URL + uri;
            } else {
                return BASE_URL + "/" + uri;
            }
        }
    </script>
    <script type="application/javascript" src="${pageContext.request.contextPath}/js/jquery.js" ></script>
    <script type="application/javascript" src="${pageContext.request.contextPath}/js/jquery.form.min.js" ></script>
    <script type="application/javascript" src="${pageContext.request.contextPath}/js/common.js"  ></script>
    <script type="application/javascript" src="${pageContext.request.contextPath}/js/easyui/jquery.easyui.min.js"  ></script>
    <script type="application/javascript" src="${pageContext.request.contextPath}/js/easyui/locale/easyui-lang-zh_CN.js"  ></script>
    <script>
        $(function(){
            $("#transASONControl_table").datagrid({
                rownumbers: true,
                pagination: true,
                singleSelect: true,
                toolbar: $('#grid_tools'),
                url: "${pageContext.request.contextPath}/json/data.json",
                method: 'get',
                title: '11111',
                width: '100%',
                height: '500',
                region: 'south',
                columns:
                        [[
                            {title:'Item Details',colspan:2},
                            {field:'itemid',title:'Item ID',rowspan:2,width:80,sortable:true},
                            {title:'1111Item Details',colspan:2},
                            {field:'productid',title:'Product ID',rowspan:2,width:80,sortable:true}
                        ],[
                            {field:'attr1',title:'Attribute',width:80},
                            {field:'status',title:'Status',width:80},
                            {field:'listprice',title:'List Price',width:80,align:'right',sortable:true},
                            {field:'unitcost',title:'Unit Cost',width:80,align:'right',sortable:true}
                        ]]
            });
            $("#transASONControl_export").on("click",function(evt){
                $.exportExcel($("#transASONControl_table"),"sheet名字","ASON控制域信息列表",urlBuild("/excel/export"));
            });
        });
    </script>
</head>
<body>
<div id="transASONControl_layout" class="easyui-layout" data-options="fit:true">
    <div data-options="region:'center',split:true">
        <div class="easyui-layout" data-options="fit:true">
            <div data-options="region:'center'">
                <div id="grid_tools">
                    <%--工具栏三个按钮--%>
                    <%--<a id="transASONControl_add" href="#" class="easyui-linkbutton" data-options="iconCls:'icon-add'" plain="true" style="vertical-align: middle;">添加</a>
                    <a id="transASONControl_save" href="#" class="easyui-linkbutton" data-options="iconCls:'icon-edit'" plain="true" style="vertical-align: middle;">编辑</a>
                    <a id="transASONControl_delete" href="#" class="easyui-linkbutton" data-options="iconCls:'icon-remove'" plain="true" style="vertical-align: middle;">删除</a>

                    --%>
                    <a id="transASONControl_export" href="#" class="easyui-linkbutton" data-options="iconCls:'icon-print'" plain="true" style="vertical-align: middle;">导出</a>
                </div>
                <%--表单栏--%>
                <table id="transASONControl_table" style="width:100%;"></table>
            </div>
        </div>
    </div>
</div>
</body>
</html>
