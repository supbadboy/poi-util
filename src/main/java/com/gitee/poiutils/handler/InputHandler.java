package com.gitee.poiutils.handler;




import com.gitee.poiutils.util.ExcelCell;

import java.util.Map;


/**
 * @author Administrator
 * @date 2020/9/9
 * @Description: TODO
 */
public class InputHandler implements DefHandler{

    @Override
    public String excute(ExcelCell excelCell, Map<String,String> param) {
        if(excelCell.getStyle() != null){
            return "<input text='text' style='"+excelCell.getStyle()+"' value=" + excelCell.getCellVal() + "></input>";
        } else {
            return "<input text='text' style='width:100%;' value=" + excelCell.getCellVal() + "></input>";
        }
    }
}
