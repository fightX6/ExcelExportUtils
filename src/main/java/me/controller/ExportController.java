package me.controller;

import me.utils.ExcelExportUtils;
import org.springframework.web.bind.annotation.*;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.util.*;

/**
 *
 */
@RestController
@RequestMapping(value = "/excel")
public class ExportController  {

    /**
     * 导出数据为excel
     * @param excelTitle 标题行json数据
     * @param request
     * @param response
     */
    @RequestMapping(value = "/export",method = RequestMethod.POST)
    public void export(String excelTitle, HttpServletRequest request, HttpServletResponse response){
        //查询行数据
        List data =  new ArrayList();
        for(int i = 0 ;i<10000;i++){
            Map<String ,Object> map = new HashMap<String, Object>();
            map.put("itemid",i+"_itemid");
            map.put("productid",i+"_productid");
            map.put("listprice",i+"_listprice");
            map.put("unitcost",i+"_unitcost");
            map.put("attr1",i+"_attr1");
            map.put("status",i+"_status");
            data.add(map);
        }
        //导出excel
        ExcelExportUtils.export(excelTitle,data,request,response);
    }
}
