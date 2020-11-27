package com.adingxiong.poiutils.handler;

import com.adingxiong.poiutils.constant.Constants;
import com.adingxiong.poiutils.interfaces.FieldName;
import com.adingxiong.poiutils.util.ExcelUtil;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;

import java.lang.reflect.Field;
import java.text.ParseException;

/**
 * @ClassName StringHandler
 * @Description TODO
 * @Author xiongchao
 * @Date 2020/11/27 16:48
 **/
public class StringHandler extends AbstractFieldParsHandler {

    @Override
    <T> void setFieldVal(Cell cell, Field filed, FieldName fieldName, StringBuffer errorMsg, T paramT) {
        String str = ExcelUtil.getCellValue(cell);
        if (StringUtils.isNotBlank(str)) {
            try {
                if (StringUtils.isNotBlank(fieldName.dateFormat())) {
                    try {
                        str = super.getDateFormat(fieldName.dateFormat()).format(Constants.simpleDateFormat.parse(str));
                    } catch (ParseException e) {
                        errorMsg.append(fieldName.value()).append("时间格式错误,");
                    }
                }
                filed.set(paramT, str);
            } catch (IllegalAccessException e) {
                errorMsg.append(fieldName.value()).append("属性值设置失败,");
            }
        }
        if ((fieldName.required()) && (StringUtils.isBlank(str))){
            errorMsg.append(fieldName.value()).append("不能为空,");
        }
    }
}
