package com.adingxiong.poiutils.handler;

import com.adingxiong.poiutils.constant.Errorcons;
import com.adingxiong.poiutils.interfaces.FieldName;
import com.adingxiong.poiutils.util.ExcelUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.springframework.util.StringUtils;

import java.lang.reflect.Field;

/**
 * @ClassName InstantHandler
 * @Description TODO
 * @Author xiongchao
 * @Date 2020/11/27 15:52
 **/
public class InstantHandler extends AbstractFieldParsHandler {
    @Override
    <T> void setFieldVal(Cell cell, Field filed, FieldName fieldName, StringBuffer errorMsg, T paramT) {
        String val = ExcelUtil.getCellValue(cell);
        if(!StringUtils.isEmpty(val)){
            String dateFormat = "yyyy-MM-dd";
            if(fieldName.dateFormat() != null) {
                dateFormat = fieldName.dateFormat();
            }
            try {
                filed.set(paramT , super.getDateFormat(dateFormat).parse(val).toInstant());
            } catch (Exception e) {
                e.printStackTrace();
                errorMsg.append(fieldName.value()).append("<").append(val).append(">").append(Errorcons.TIME_TYPE);
            }
        }
        if(fieldName.required() && StringUtils.isEmpty(val)){
            errorMsg.append(fieldName.value()).append(Errorcons.NOT_EMPTY);
        }
    }
}
