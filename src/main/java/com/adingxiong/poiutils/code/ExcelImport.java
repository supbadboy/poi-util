package com.adingxiong.poiutils.code;


import com.adingxiong.poiutils.constant.Constants;
import com.adingxiong.poiutils.constant.Errorcons;
import com.adingxiong.poiutils.handler.FieldParsHandler;
import com.adingxiong.poiutils.interfaces.FieldName;
import com.adingxiong.poiutils.util.ExcelUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.springframework.util.Assert;
import org.springframework.util.StringUtils;

import java.lang.reflect.Field;
import java.text.NumberFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import static com.adingxiong.poiutils.constant.Errorcons.*;

/**
 * @ClassName ExcelImport
 * @Description   Excel导出
 * @Author xiongchao
 * @Date 2020/11/26 11:04
 **/
public class ExcelImport {

    /**
     * 标题所在行号 默认为 0
     */
    private Integer rowNum = Integer.valueOf(0);

    /**
     * 是否格式话标题  去掉 空格  换行符
     */
    private Boolean formatTitle = Boolean.valueOf(true);

    /**
     * 指定存放错误日志的字段名
     */
    private String fieldError =  null ;

    /**
     * 指定记录表格行数的字段名
     */
    private String fieldRows =  null;

    private NumberFormat format = null ;

    private ExcelImport(){
        this.format = NumberFormat.getInstance();
        this.format.setMaximumFractionDigits(2);
        this.format.setGroupingUsed(true);
    }

    public static ExcelImport getInstance (){
        return new ExcelImport();
    }

    public ExcelImport setRowNum(Integer rowNum) {
        Assert.notNull(rowNum , PARM_EMPTY);
        this.rowNum = rowNum;
        return this;
    }

    public ExcelImport isFormatTitle(Boolean formatTitle) {
        Assert.notNull(formatTitle , PARM_EMPTY);
        this.formatTitle = formatTitle;
        return this;
    }

    public ExcelImport setFieldError(String fieldError) {
        Assert.notNull(fieldError , PARM_EMPTY);
        this.fieldError = fieldError;
        return this;
    }

    public ExcelImport setFieldRows(String fieldRows) {
        Assert.notNull(fieldRows , PARM_EMPTY);
        this.fieldRows = fieldRows;
        return this;
    }

    public synchronized <T> List<T> transformation(Sheet sheet , Class<T> clazz){
        List result = new ArrayList();
        Assert.notNull(sheet ,ST_EMPTY);
        Assert.notNull(clazz ,CLASSS_EMPTY);
        try {
            Field[] fields = clazz.getDeclaredFields();
            Map titleMap  = ExcelUtil.rowMap(sheet.getRow(this.rowNum),fields ,this.formatTitle);
            if(titleMap == null || titleMap.isEmpty()){
                throw new NullPointerException(EXCEL_TITLE_EMPTY);
            }
            Integer rowNum = Integer.valueOf(sheet.getLastRowNum() + 1);
            Row row = null ;
            Cell cell = null;
            Object t = null ;
            for (int i = this.rowNum.intValue() +1 ; i < rowNum ; i++) {
                t = clazz.newInstance();
                row = sheet.getRow(i);
                StringBuffer errorMsg = new StringBuffer();
                for(Field field :fields) {
                    if(field.isAnnotationPresent(FieldName.class) && titleMap.containsKey(field.getName())){
                        FieldName fieldName = field.getAnnotation(FieldName.class);
                        int index = (int) titleMap.get(field.getName());
                        cell = row.getCell(index);
                        if(cell != null) {
                            String fileType = field.getType().getSimpleName();
                            field.setAccessible(true);
                            FieldParsHandler fieldParsHandler = null;
                            if(FieldParsHandler.handles.containsKey(fileType + Constants.HEAD_SUFFIX)){
                                fieldParsHandler = FieldParsHandler.handles.get(fileType + Constants.HEAD_SUFFIX);
                            }else {
                                Class handler = Class.forName("com.adingxiong.poiutils.handler." + fileType + Constants.HEAD_SUFFIX);
                                fieldParsHandler = (FieldParsHandler) handler.newInstance();
                                FieldParsHandler.handles.put(fileType + Constants.HEAD_SUFFIX , fieldParsHandler);
                            }
                            Assert.notNull(fieldParsHandler , Errorcons.INSTANCE_ERROR);
                            fieldParsHandler.execute(cell,field,fieldName,errorMsg,t);
                        } else if  (fieldName.required()){
                            errorMsg.append(fieldName.value()).append(Errorcons.NOT_EMPTY);
                        }
                    }
                    if(!StringUtils.isEmpty(this.fieldError)){
                        Field fd = clazz.getDeclaredField(this.fieldError);
                        fd.setAccessible(true);
                        if(errorMsg.length() != 0) {
                            fd.set(t , errorMsg.toString().substring(0 ,errorMsg.toString().length() - 1));
                        }
                    }
                    if (!StringUtils.isEmpty(this.fieldRows)) {
                        Field rows = clazz.getDeclaredField(this.fieldRows);
                        rows.setAccessible(true);
                        rows.set(t, Integer.valueOf(i));
                    }
                }
                result.add(t);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return result;
    }

}
