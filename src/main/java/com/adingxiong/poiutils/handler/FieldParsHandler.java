package com.adingxiong.poiutils.handler;

import com.adingxiong.poiutils.interfaces.FieldName;
import org.apache.poi.ss.usermodel.Cell;

import java.lang.reflect.Field;
import java.util.concurrent.ConcurrentHashMap;

/**
 * @ClassName FieldParsHandler
 * @Description TODO
 * @Author xiongchao
 * @Date 2020/11/27 15:18
 **/
public abstract interface FieldParsHandler {

    public static final ConcurrentHashMap<String, FieldParsHandler> handles = new ConcurrentHashMap();

    public abstract <T> void execute (Cell cell , Field filed , FieldName fieldName , StringBuffer errorMsg , T param);
}
