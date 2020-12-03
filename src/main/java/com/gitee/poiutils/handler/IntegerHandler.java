package com.gitee.poiutils.handler;

import com.gitee.poiutils.constant.Errorcons;
import com.gitee.poiutils.interfaces.FieldName;
import com.gitee.poiutils.util.ExcelUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.springframework.util.StringUtils;

import java.lang.reflect.Field;
import java.util.regex.Pattern;

/**
 * ClassName IntegerHandler
 * Description TODO
 * @author xiongchao
 * Date 2020/11/27 15:46
 **/
public class IntegerHandler extends AbstractFieldParsHandler {
    @Override
    <T> void setFieldVal(Cell cell, Field filed, FieldName fieldName, StringBuffer errorMsg, T paramT) {
        String val = ExcelUtil.getCellValue(cell);
        if(!StringUtils.isEmpty(val)){
            if(!Pattern.matches(com.gitee.poiutils.constant.Pattern.ISNUM , val)){
                errorMsg.append(fieldName.value()).append("<").append(val).append(">").append(Errorcons.PARAM_TYPE_ERROR);
            }else{
                try {
                    filed.set(paramT,Integer.parseInt(val));
                } catch (IllegalAccessException e) {
                    e.printStackTrace();
                    errorMsg.append(fieldName.value()).append(Errorcons.SET_ERROR);
                }
            }
        }
        if(fieldName.required() && StringUtils.isEmpty(val)){
            errorMsg.append(fieldName.value()).append(Errorcons.NOT_EMPTY);
        }
    }
}
