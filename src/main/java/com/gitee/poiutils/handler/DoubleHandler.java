package com.gitee.poiutils.handler;

import com.gitee.poiutils.constant.Errorcons;
import com.gitee.poiutils.interfaces.FieldName;
import com.gitee.poiutils.util.ExcelUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.springframework.util.StringUtils;

import java.lang.reflect.Field;
import java.util.regex.Pattern;

import static com.gitee.poiutils.constant.Errorcons.PARAM_TYPE_ERROR;

/**
 * ClassName DoubleHandler
 * Description TODO
 * @author xiongchao
 * Date 2020/11/27 15:36
 **/
public class DoubleHandler extends AbstractFieldParsHandler {
    @Override
    <T> void setFieldVal(Cell cell, Field filed, FieldName fieldName, StringBuffer errorMsg, T paramT) {
        String db = ExcelUtil.getCellValue(cell);
        if(!StringUtils.isEmpty(db)){
            if(!Pattern.matches(com.gitee.poiutils.constant.Pattern.ISDOUBLE , db)){
                errorMsg.append(fieldName.value()).append("<").append(db).append(">").append(PARAM_TYPE_ERROR);
            }else {
                try {
                    filed.set(paramT ,Double.valueOf(db));
                } catch (IllegalAccessException e) {
                    e.printStackTrace();
                    errorMsg.append(fieldName.value()).append(Errorcons.SET_ERROR);
                }
            }
        }
        if(fieldName.required() && StringUtils.isEmpty(db)){
            errorMsg.append(fieldName.value()).append(Errorcons.NOT_EMPTY);
        }
    }
}
