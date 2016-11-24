package me.utils;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import org.apache.commons.beanutils.BeanMap;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFCellUtil;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.util.CellRangeAddress;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.OutputStream;
import java.sql.Timestamp;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * 公用 excel easyui的datagrid 表格导出
 * Created by xf on 2016/10/9.
 */
public class ExcelExportUtils {
    /**
     *  公共导出excel方法
     * @param excelTitleData 导出excel的标题json数据
     *                       格式为{
    TitleInfo: $(...).datagrid("options").columns,
    SheetName: "sheet名字",
    SaveName: "表单excel文件名称"
    }
    特别说明 列属性上可加dataSelect属性来 格式化  显示值    例如定义  dataSelect:"1#正常|2#已删除|3#注销"
     * @param data 导出excel的行数据
     * @param request HttpServletRequest
     * @param response HttpServletResponse
     * @return
     */
    public static void export(String  excelTitleData,List<? extends Object> data,HttpServletRequest request,HttpServletResponse response){
        List<Map<String,Object>>  rowsData = new ArrayList<Map<String, Object>>();
        if(data != null && data.size() > 0 ) {
            for (Object obj : data) {
                if(! (obj instanceof  Map)) {
                    Map map = new BeanMap(obj);
                    rowsData.add(map);
                }else{
                    rowsData.add((Map<String,Object>)obj);
                }
            }
        }
        //创建工作簿
        HSSFWorkbook workBook = new HSSFWorkbook();
        if( excelTitleData == null || "".equals( excelTitleData)){
            return ;
        }
        //将字符串数据转换为json数据
        JSONObject obj = JSONArray.parseObject(excelTitleData);
        String workBookName = "";
        if(obj != null) {
            workBookName = obj.getString("SaveName");
            String sheetName = obj.getString("SheetName");
            HSSFSheet sheet = workBook.createSheet(sheetName);
            //行开始序号
            int rowNum = 0;
            //创建标题行
            Object[] props = buildHeadRow(workBook,sheet,obj,rowNum);
            rowNum = (Integer)props[0];
            Object[] cells = (Object[])props[1];

            //创建数据行
            buildDataRow(workBook,sheet,rowsData,rowNum,cells);
        }
        out(workBook,workBookName,request,response);
    }
    /**
     *  公共导出excel方法
     * @param excelTitleData 导出excel的标题json数据
     *                       格式为{
    TitleInfo: $(...).datagrid("options").columns,
    SheetName: "sheet名字",
    SaveName: "表单excel文件名称"
    }
    特别说明 列属性上可加dataSelect属性来 格式化  显示值    例如定义  dataSelect:"1#正常|2#已删除|3#注销"
     * @param data 导出excel的行数据
     * @param request HttpServletRequest
     * @param response HttpServletResponse
     * @return
     */
    public static Object[] exportDiy(String excelTitleData,List<? extends Object> data,HttpServletRequest request,HttpServletResponse response){
        List<Map<String,Object>>  rowsData = new ArrayList<Map<String, Object>>();
        if(data != null && data.size() > 0 ) {
            for (Object obj : data) {
                if(! (obj instanceof  Map)) {
                    Map map = new BeanMap(obj);
                    rowsData.add(map);
                }else{
                    rowsData.add((Map<String,Object>)obj);
                }
            }
        }
        //创建工作簿
        HSSFWorkbook workBook = new HSSFWorkbook();
        if(excelTitleData == null || "".equals(excelTitleData)){
            return null;
        }
        //将字符串数据转换为json数据
        JSONObject obj = JSONArray.parseObject(excelTitleData);
        String workBookName = "";
        Object[] cells=null;
        String sheetName = "";
        if(obj != null) {
            workBookName = obj.getString("SaveName");
            sheetName = obj.getString("SheetName");
            HSSFSheet sheet = workBook.createSheet(sheetName);
            //行开始序号
            int rowNum = 0;
            //创建标题行
            Object[] props = buildHeadRow(workBook,sheet,obj,rowNum);
            rowNum = (Integer)props[0];
            cells = (Object[])props[1];

            //创建数据行
            buildDataRow(workBook,sheet,rowsData,rowNum,cells);
        }
        return new Object[]{workBook,workBookName,cells,sheetName};
    }

    public static void out(HSSFWorkbook workBook,String workBookName ,HttpServletRequest request,HttpServletResponse response){
        if(workBook !=null){
            OutputStream out = null;
            try{
                String fileName = workBookName + ".xls";
                String headStr = "attachment; filename=\"" + processFileName(request,fileName) + "\"";
                response.setContentType("APPLICATION/OCTET-STREAM");
                response.setHeader("Content-Disposition", headStr);
                out = response.getOutputStream();
                workBook.write(out);
                out.flush();
            }catch (IOException e){
                e.printStackTrace();
            }finally {
                try{
                    if(out != null) {
                        out.close();
                    }
                }catch (IOException e){
                    e.printStackTrace();
                }
            }
        }
    }
    /**
     * 构建标题行 返回表头属性
     * @param workBook excel工作簿
     * @param sheet  excel工作空间
     * @param obj  标题行数据
     * @param rowNum
     * @return Object[3]  0:标题行的行数  1:与数据对应的顺序列 List<JSONObject>
     */
    private static Object[] buildHeadRow(HSSFWorkbook workBook,HSSFSheet sheet,JSONObject obj,int rowNum){
        Object[] rowInfo = new Object[2];
        //定义数据行需要按顺序取得哪些字段列
        Object[] cells = null;
        //获取整个标题行的数据
        JSONArray columns = obj.getJSONArray("TitleInfo");

        HSSFCellStyle style = setColumnTopStyle(workBook);
        if( columns != null && columns.size() > 0 ) {
            int colCount = 0;//获取总列数
            if (true) {
                JSONArray cols = columns.getJSONArray(0);
                for(Object c:cols) {
                    JSONObject cell = (JSONObject)c;
                    Integer colspan = cell.getInteger("colspan");
                    colspan = colspan == null ? 1 : colspan;
                    colCount += colspan;
                }
                cells = new Object[colCount];
            }
            for (int i = 0; i < columns.size() ; i++ ){
                //获取标题行的一行的数据
                JSONArray cols = columns.getJSONArray(i);
                if(cols != null && cols.size() > 0){
                    HSSFRow row = sheet.createRow(rowNum++);
                    int colNum = 0;
                    for(int j = 0; j < cols.size(); j++){
                        //获取标题行的一行的一个单元格数据
                        JSONObject cell = cols.getJSONObject(j);
                        if(cell != null && (cell.getBoolean("hidden") == null || cell.getBoolean("hidden") != null && !cell.getBoolean("hidden") )
                                && (cell.getBoolean("checkbox") == null || cell.getBoolean("checkbox") != null && !cell.getBoolean("checkbox")) ){
                            Integer rowspan = cell.getInteger("rowspan");
                            rowspan = rowspan==null?1:rowspan;
                            Integer colspan = cell.getInteger("colspan");
                            colspan = colspan==null?1:colspan;
                            String title = cell.getString("title");
                            String align = cell.getString("align");
                            // 获得一个 sheet 中合并单元格的数量
                            int sheetmergerCount = sheet.getNumMergedRegions();
                            // 遍历合并单元格
                            for (int c = 0; c < sheetmergerCount; c++) {
                                // 获得合并单元格
                                CellRangeAddress ca = sheet.getMergedRegion(c);
                                // 获得合并单元格的起始行, 结束行, 起始列, 结束列
                                int firstC = ca.getFirstColumn();
                                int firstR = ca.getFirstRow();
                                int lastC = ca.getLastColumn();
                                int lastR = ca.getLastRow();
                                if( colNum <= lastC && colNum >= firstC && i <= lastR && i >= firstR){
                                    colNum = lastC + 1;//
                                }
                            }
                            //若单元格为最后一行的单元格则记录到需要对应数据的集合中
                            if( (i + rowspan )  == columns.size() ){
                                if(cells[colNum] == null ) {
                                    cells[colNum] = cell;
                                }else{
                                    for(int c = colNum ;c < cells.length ; c++){
                                        if(cells[c] == null){
                                            cells[c] = cell;
                                            break;
                                        }
                                    }
                                }
                            }
                            //创建单元格
                            HSSFCell rowCell = row.createCell(colNum);
                            //设置单元格内容
                            rowCell.setCellValue(title);
                            //合并单元格
                            if((i+rowspan-1) != i || colNum != (colNum+colspan-1)) {
                                CellRangeAddress region = new CellRangeAddress(i, i + rowspan - 1, colNum, colNum + colspan - 1);
                                sheet.addMergedRegion(region);
                            }
                            //设置样式
                            sheet.setColumnWidth(colNum,getStringByteLength(title)*256*5);
                            setStyleAlign("center",style);
                            rowCell.setCellStyle(style);
                            colNum++;
                        }
                    }
                }
            }
            //设置合并的单元格样式
            // 遍历合并单元格
            for (int c = 0; c < sheet.getNumMergedRegions(); c++) {
                // 获得合并单元格
                setStyleRegion(sheet, sheet.getMergedRegion(c), style);
            }
        }
        rowInfo[0] = rowNum;//标题行的行数
        rowInfo[1] = cells;//需要取值的列字段对象Cell
        return  rowInfo;
    }

    /**
     *  构建数据行
     * @param workBook excel工作簿
     * @param sheet  excel工作空间
     * @param rowsData 行数据
     * @param rowNum 行号
     * @param cells 需要添加数据的列字段数据
     */
    private static void buildDataRow(HSSFWorkbook workBook,HSSFSheet sheet,List<Map<String,Object>> rowsData,int rowNum,Object[] cells){
        HSSFCellStyle redStyle = setStyle(workBook,"red");
        HSSFCellStyle blueStyle = setStyle(workBook,"blue");
        HSSFCellStyle blackStyle = setStyle(workBook,null);
        HSSFCellStyle style = blackStyle;
        if(rowsData != null && rowsData.size() > 0 ){
            for (Map row :rowsData){
                if( cells  != null && cells.length > 0 ) {
                    HSSFRow hssfRow = sheet.createRow(rowNum);
                    int colNum = 0;
                    //设置行颜色 红色  蓝色 默认  黑色
                    String color = row.get("color") != null && !row.get("color").toString().equalsIgnoreCase("")?row.get("color").toString():"";
                    if(!color.equalsIgnoreCase("")){
                        // 设置字体
                        if(color.equalsIgnoreCase("red")){
                            style = redStyle;
                        }else if(color.equalsIgnoreCase("blue")){
                            style = blueStyle;
                        }
                    }
                    for (Object obj : cells) {
                        JSONObject cell = (JSONObject)obj;
                        if(cell != null) {
                            Integer colspan = cell.getInteger("colspan");
                            colspan = colspan == null ? 1 : colspan;
                            String field = cell.getString("field");
                            String dataSelect = cell.getString("dataSelect");//
                            String align = cell.getString("align");
                            String timeFormat = cell.getString("timeFormat");
                            String extValue = cell.getString("extValue");//额外的拼接字段值
                            //创建单元格
                            HSSFCell hssfCell = hssfRow.createCell(colNum);
                            //设置单元格内容
                            Object value = row.get(field);
                            if (dataSelect != null && !"".equals(dataSelect)) {
                                hssfCell.setCellValue(value != null ? getSelectValue(dataSelect, value.toString()) : "");
                            } else {
                                if (value instanceof Timestamp || value instanceof Date) {
                                    SimpleDateFormat format = new SimpleDateFormat(timeFormat != null && !"".equalsIgnoreCase(timeFormat) ? timeFormat : "yyyy-MM-dd");
                                    value = format.format(value);
                                }
                                if(field.toLowerCase().contains("password")){
                                    hssfCell.setCellValue("******");
                                }else{
                                    hssfCell.setCellValue(value != null ? value.toString() : "");
                                }
                            }
                            //添加额外字段
                            if(extValue != null && !extValue.equals("")){
                                hssfCell.setCellValue(hssfCell.getStringCellValue() + extValue);
                            }
                            //合并单元格
                            if (colNum != (colNum + colspan - 1)) {
                                CellRangeAddress region = new CellRangeAddress(rowNum, rowNum, colNum, colNum + colspan - 1);
                                sheet.addMergedRegion(region);
                            }
                            //设置样式
                            setStyleAlign("center", style);
                            hssfCell.setCellStyle(style);
                            colNum += colspan;
                        }
                    }
                    //一行完成 颜色换为 黑色
                    style = blackStyle;
                    rowNum++;
                }
            }
        }
    }

    /**
     * 获取格式化字段值
     * @param dataSelect 格式化规则  例如  1#正常|2#已删除|3#注销
     * @param cellValue  需要格式化的值
     * @return
     */
    private static String getSelectValue(String dataSelect,String cellValue){
        Map<String,String> select = new HashMap<String, String>();
        if(dataSelect != null && !"".equals(dataSelect)) {
            String[] items = dataSelect.split("\\|");
            if(items != null && items.length > 0) {
                for (String item : items) {
                    String[] item_c = item.split("#");
                    if(item_c != null && item_c.length == 2) {
                        if(item_c[0] != null && !"".equals(item_c[0]) )
                            if(item_c[1].equalsIgnoreCase("null")){
                                select.put(item_c[0],null);
                            }else {
                                select.put(item_c[0], item_c[1]);
                            }
                    }
                }
            }
        }
        if(cellValue != null && !cellValue.equals("")){
            if(select.get("anon") != null){//设置为anon的 表示  不满足其他条件的统一值
                if(select.get(cellValue) == null){
                    return select.get("anon");
                }
            }
            return select.get(cellValue);
        }
        return "";
    }

    /**
     * 设置合并的单元格样式
     * @param sheet
     * @param region
     * @param cs
     */
    public static void setStyleRegion(HSSFSheet sheet, CellRangeAddress region, HSSFCellStyle cs) {
        for (int i = region.getFirstRow(); i <= region.getLastRow(); i++) {
            HSSFRow row = HSSFCellUtil.getRow(i, sheet);
            for (int j = region.getFirstColumn(); j <= region.getLastColumn(); j++) {
                HSSFCell cell = HSSFCellUtil.getCell(row, (short) j);
                cell.setCellStyle(cs);
            }
        }
    }
    /**
     * 设置位置样式的
     * @param align
     * @param style
     */
    public static void setStyleAlign(String align,HSSFCellStyle style){
        if(align != null && !align.equals("")){
            // 设置居左
            if(align.equalsIgnoreCase("left")){ style.setAlignment(HSSFCellStyle.ALIGN_LEFT);}
            // 设置居中
            else if(align.equalsIgnoreCase("center")){ style.setAlignment(HSSFCellStyle.ALIGN_CENTER); }
            // 设置居右
            else if(align.equalsIgnoreCase("right")){ style.setAlignment(HSSFCellStyle.ALIGN_RIGHT); }
            else{
                style.setAlignment(HSSFCellStyle.ALIGN_CENTER);// 设置居中
            }
        }else{
            style.setAlignment(HSSFCellStyle.ALIGN_CENTER);// 设置居中
        }
    }
    /*
     * 列头单元格样式
     */
    public static HSSFCellStyle setColumnTopStyle(HSSFWorkbook workbook ) {
        // 设置字体
        HSSFFont font = workbook.createFont();
        //设置字体大小
        font.setFontHeightInPoints((short)11);
        //字体加粗
        font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
        //设置字体名字
        font.setFontName("Courier New");
        font.setColor(HSSFFont.COLOR_NORMAL);
//      String align = cellData.getString("align");
        //设置样式;
        HSSFCellStyle cell_Style = workbook.createCellStyle();// 设置样式
        //在样式用应用设置的字体;
        cell_Style.setFont(font);
        cell_Style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);// 垂直对齐居中
        //设置底边框;
        cell_Style.setBorderBottom(HSSFCellStyle.BORDER_THICK);//粗的
        //设置底边框颜色;
        cell_Style.setBottomBorderColor(HSSFColor.BLACK.index);
        //设置左边框;
        cell_Style.setBorderLeft(HSSFCellStyle.BORDER_THICK);
        //设置左边框颜色;
        cell_Style.setLeftBorderColor(HSSFColor.BLACK.index);
        //设置右边框;
        cell_Style.setBorderRight(HSSFCellStyle.BORDER_THICK);
        //设置右边框颜色;
        cell_Style.setRightBorderColor(HSSFColor.BLACK.index);
        //设置顶边框;
        cell_Style.setBorderTop(HSSFCellStyle.BORDER_THICK);
        //设置顶边框颜色;
        cell_Style.setTopBorderColor(HSSFColor.BLACK.index);
        //在样式用应用设置的字体;
        cell_Style.setFont(font);
        //设置自动换行;
        cell_Style.setWrapText(false);
        //设置水平对齐的样式为居中对齐;
        cell_Style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        //设置垂直对齐的样式为居中对齐;
        cell_Style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
        return cell_Style;
    }
    /**
     *  列数据信息单元格样式
     * @param workbook
     * @param //cell
     * @param //cellData
     * @return
     */
    public static HSSFCellStyle setStyle(HSSFWorkbook workbook,String color){
        // 设置字体
        HSSFFont font = workbook.createFont();
        //设置字体大小
        //font.setFontHeightInPoints((short)10);
        //字体加粗
        //font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
        //设置字体名字
        //设置行颜色 红色  蓝色 默认  黑色
        font.setColor(Font.COLOR_NORMAL);
        if(color != null && !color.equalsIgnoreCase("")){
            // 设置字体
            if(color.equalsIgnoreCase("red")){
                font.setColor(HSSFColor.RED.index);
            }else if(color.equalsIgnoreCase("blue")){
                font.setColor(HSSFColor.BLUE.index);
            }
        }
        font.setFontName("Courier New");
//        String align = cellData.getString("align");
        HSSFCellStyle cell_Style = workbook.createCellStyle();// 设置样式
        //在样式用应用设置的字体;
        cell_Style.setFont(font);
        cell_Style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);// 垂直对齐居中
        cell_Style.setWrapText(false); // 设置为不自动换行
        //设置底边框;
        cell_Style.setBorderBottom(HSSFCellStyle.BORDER_THIN);//细的
        //设置底边框颜色;
        cell_Style.setBottomBorderColor(HSSFColor.BLACK.index);
        //设置左边框;
        cell_Style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        //设置左边框颜色;
        cell_Style.setLeftBorderColor(HSSFColor.BLACK.index);
        //设置右边框;
        cell_Style.setBorderRight(HSSFCellStyle.BORDER_THIN);
        //设置右边框颜色;
        cell_Style.setRightBorderColor(HSSFColor.BLACK.index);
        //设置顶边框;
        cell_Style.setBorderTop(HSSFCellStyle.BORDER_THIN);
        //设置顶边框颜色;
        cell_Style.setTopBorderColor(HSSFColor.BLACK.index);
        return cell_Style;
    }

    /**
     *
     * @Title: processFileName
     *
     * @Description: ie,chrom,firfox下处理文件名显示乱码
     */
    public static String processFileName(HttpServletRequest request, String fileNames) {
        String filename = "";
        try {
            String agent = request.getHeader("USER-AGENT");
            if (null != agent && -1 != agent.indexOf("MSIE") || null != agent && -1 != agent.indexOf("Trident")|| null != agent && -1 != agent.indexOf("Edge")) {// ie
                filename = java.net.URLEncoder.encode(fileNames, "UTF8");
            } else if (null != agent && -1 != agent.indexOf("Mozilla")) {// 火狐,chrome等
                filename = new String(fileNames.getBytes("UTF-8"), "iso-8859-1");
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return filename;
    }
    public static int getStringByteLength(String s) {
        int length = 0;
        for(int i = 0; i < s.length(); i++)
        {
            int ascii = Character.codePointAt(s, i);
            if(ascii >= 0 && ascii <=255)
                length++;
            else
                length += 2;

        }
        return length;
    }
}