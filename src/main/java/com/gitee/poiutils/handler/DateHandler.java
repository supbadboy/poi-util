package com.gitee.poiutils.handler;

import com.gitee.poiutils.constant.Errorcons;
import com.gitee.poiutils.interfaces.FieldName;
import com.gitee.poiutils.util.ExcelUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.springframework.util.StringUtils;

import java.lang.reflect.Field;

/**
 * ClassName DateHandler
 * Description TODO
 * @author xiongchao
 * Date 2020/11/27 15:27
 **/
public class DateHandler extends AbstractFieldParsHandler {
    @Override
    <T> void setFieldVal(Cell paramCell, Field paramField, FieldName paramFieldName, StringBuffer paramStringBuffer, T paramT) {
        String date = ExcelUtil.getCellValue(paramCell);
        if(!StringUtils.isEmpty(date)){
            String dateFormat = "yyyy-MM-dd" ;
            if(!StringUtils.isEmpty(paramFieldName.dateFormat())){
                dateFormat = paramFieldName.dateFormat();
            }
            try {
                paramField.set(paramT , super.getDateFormat(dateFormat).parse(date));
            } catch (Exception e) {
                e.printStackTrace();
                paramStringBuffer.append(paramFieldName.value()).append("<").append(date).append(">").append(Errorcons.TIME_TYPE);
            }
        }
        if(paramFieldName.required() && StringUtils.isEmpty(date)){
            paramStringBuffer.append(paramFieldName.value()).append(Errorcons.NOT_EMPTY);
        }
    }
}
