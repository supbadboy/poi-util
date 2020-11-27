package com.adingxiong.poiutils.handler;

import com.adingxiong.poiutils.constant.Errorcons;
import com.adingxiong.poiutils.interfaces.FieldName;
import com.adingxiong.poiutils.util.ExcelUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.springframework.util.StringUtils;

import java.lang.reflect.Field;
import java.util.regex.Pattern;

import static com.adingxiong.poiutils.constant.Pattern.ISNUM;

/**
 * @ClassName IntegerHandler
 * @Description TODO
 * @Author xiongchao
 * @Date 2020/11/27 15:46
 **/
public class IntegerHandler extends AbstractFieldParsHandler {
    @Override
    <T> void setFieldVal(Cell cell, Field filed, FieldName fieldName, StringBuffer errorMsg, T paramT) {
        String val = ExcelUtil.getCellValue(cell);
        if(!StringUtils.isEmpty(val)){
            if(!Pattern.matches(ISNUM , val)){
                errorMsg.append(fieldName.value()).append(Errorcons.PARAM_TYPE_ERROR);
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
