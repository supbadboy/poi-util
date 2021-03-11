package com.gitee.poiutils.handler;




import com.gitee.poiutils.util.ExcelCell;

import java.util.*;

/**
 * @author Administrator
 * @date 2020/9/9
 * @Description: TODO
 */
public class SelectHandler implements DefHandler{

    @Override
    public String excute(ExcelCell excelCell, Map<String,String> param) {
        try {
            List<Map<String, String>> options = getOptions(excelCell.getEnums());
            StringBuffer stringBuffer = new StringBuffer();

            if(excelCell.getStyle() != null){
                stringBuffer.append("<select style='"+excelCell.getStyle()+"' value=" + excelCell.getCellVal() + ">");
            } else {
                stringBuffer.append("<select style='width:100%' value=" + excelCell.getCellVal() + ">");
            }
            for(Map<String,String> map : options){
                stringBuffer.append("<option value='"+map.get("val")+"'>"+map.get("name")+"</option>");
            }
            stringBuffer.append("</select>");
            return stringBuffer.toString();
        } catch (Exception e) {
            return "<span style='color:red'>配置错误<span>";
        }
    }

    private List<Map<String,String>> getOptions(String enums){
        List<Map<String,String>> list = new ArrayList<Map<String, String>>();
        if(enums != null){
            Arrays.asList(enums.split(";")).forEach(item->{
                Map<String,String> temp = new HashMap<>();
                String[] split = item.split(":");
                temp.put("val",split[0]);
                temp.put("name",split[1]);
                list.add(temp);
            });
        }
        return list;
    }
}
