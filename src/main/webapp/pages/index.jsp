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
            $("#datagrid_table").datagrid({
                rownumbers: true,
                pagination: true,
                singleSelect: true,
                toolbar: $('#grid_tools'),
                url: "${pageContext.request.contextPath}/json/data.json",
                method: 'get',
                title: '多行表单',
                width: '100%',
                height: '500',
                region: 'south',
                columns:
                        [[
                            {title:'Item Details',colspan:3},
                            {title:'1111Item Details',colspan:2},
                            {field:'productid',title:'Product ID',rowspan:2,width:80,sortable:true}
                        ],[
                            {field:'attr1',title:'Attribute',width:80},
                            {field:'status',title:'Status',width:80},
                            {field:'itemid',title:'Item ID',width:80,sortable:true},
                            {field:'listprice',title:'List Price',width:80,align:'right',sortable:true},
                            {field:'unitcost',title:'Unit Cost',width:80,align:'right',sortable:true}
                        ]]
            });
            $("#export").on("click",function(evt){
                $.exportExcel($("#datagrid_table"),"多行表单sheet名字","多行表单excel文件名称",urlBuild("/excel/export"));
            });
            $("#datagrid_table_single").datagrid({
                rownumbers: true,
                pagination: true,
                singleSelect: true,
                toolbar: $('#grid_tools_single'),
                url: "${pageContext.request.contextPath}/json/data.json",
                method: 'get',
                title: '单行表单',
                width: '100%',
                height: '500',
                region: 'south',
                columns:
                        [[
                            {field:'itemid',title:'Item ID',width:80,sortable:true},
                            {field:'productid',title:'Product ID',width:80,sortable:true},
                            {field:'attr1',title:'Attribute',width:80},
                            {field:'status',title:'Status',width:80},
                            {field:'listprice',title:'List Price',width:80,align:'right',sortable:true},
                            {field:'unitcost',title:'Unit Cost',width:80,align:'right',sortable:true}
                        ]]
            });
            $("#export_single").on("click",function(evt){
                $.exportExcel($("#datagrid_table_single"),"单行表单sheet名字","单行表单excel文件名称",urlBuild("/excel/export"));
            });
        });
    </script>
</head>
<body>
<div id="layout" class="easyui-layout" data-options="fit:true">
    <div data-options="region:'center',split:true">
        <div class="easyui-layout" data-options="fit:true">
            <div data-options="region:'center'">
                <%-- 单行表单 --%>
                <div id="grid_tools_single">
                    <%--工具栏按钮--%>
                    <a id="export_single" href="#" class="easyui-linkbutton" data-options="iconCls:'icon-print'" plain="true" style="vertical-align: middle;">导出</a>
                </div>
                <%--表单栏--%>
                <table id="datagrid_table_single" style="width:100%;"></table>
                <%-- 多行表单 --%>
                <div id="grid_tools">
                    <%--工具栏按钮--%>
                    <a id="export" href="#" class="easyui-linkbutton" data-options="iconCls:'icon-print'" plain="true" style="vertical-align: middle;">导出</a>
                </div>
                <table id="datagrid_table" style="width:100%;"></table>
            </div>
        </div>
    </div>
</div>
</body>
</html>
