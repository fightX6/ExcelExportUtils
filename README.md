# ExcelExportUtils
easyui datagrid标题格式  公共导出
Datagrid的列配置中，导出时不会导出hidden属性为true的列和checkbox属性为true的列，
注意 列配置中 可以配置dataSelect属性，，格式按照xx#yy|xx#yy|...  xx表示该字段的值，yy表示要显示在excel单元格上的值，特殊的xx=anno时单元格的值不为其他xx的值的时候都会转化为这个xx=anno对应的yy的值。 例如：dataSelect:"0#否|1#是"
              可以配置timeFormat属性，格式按照java的格式化时间格式写 例如"yyyy-MM-dd"
              可以配置extValue属性 ，需要增加的额外要拼接在单元格值之后的值
     后台数据中 字段可增加color字段 来控制行颜色  值有red和blue 不区分大小写 默认为黑色


