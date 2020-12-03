package com.gitee.poiutils.handler;

import com.gitee.poiutils.interfaces.FieldName;
import org.apache.poi.ss.usermodel.Cell;
import org.springframework.util.Assert;

import java.lang.reflect.Field;
import java.text.SimpleDateFormat;

import static com.gitee.poiutils.constant.Errorcons.PARM_EMPTY;

/**
 * ClassName AbstractFieldParsHandler
 * Description TODO
 * @author xiongchao
 * Date 2020/11/27 15:22
 **/
public abstract class AbstractFieldParsHandler implements FieldParsHandler{

    private ThreadLocal<SimpleDateFormat> simpleDateFormat = new ThreadLocal<>();

    private ThreadLocal<String> dateFormat = new ThreadLocal<>();

    @Override
    public <T> void execute(Cell cell, Field filed, FieldName fieldName, StringBuffer errorMsg, T param) {
        setFieldVal(cell,filed,fieldName,errorMsg,param);
    }

    abstract <T> void setFieldVal(Cell cell, Field filed, FieldName fieldName, StringBuffer errorMsg, T paramT);


    public SimpleDateFormat getDateFormat(String dateFormat) {
        Assert.notNull(dateFormat, PARM_EMPTY);
        if ((this.simpleDateFormat.get() != null) && (dateFormat.equals(this.dateFormat.get()))) {
            return (SimpleDateFormat)this.simpleDateFormat.get();
        }
        this.dateFormat.set(dateFormat);
        this.simpleDateFormat.set(new SimpleDateFormat(dateFormat));
        return (SimpleDateFormat)this.simpleDateFormat.get();
    }
}
