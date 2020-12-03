package com.gitee.poiutils.util;


import com.gitee.poiutils.constant.Constants;
import com.gitee.poiutils.interfaces.FieldName;
import org.springframework.util.Assert;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * ClassName ExcelUtil
 * Description   操作实体类的工具  用于读取指定注解的实体类的注解内容和 注解列名
 * @author xiongchao
 * Date 2020/11/26 13:55
 **/
public class ClassUtils {

    /**
     * 获取实体类  有指定注解的 类名和注解名
     * @param  instance  需要获取的对象
     * @return 返回对象字段和注释信息
     */
    public static Map<String,List> getDeclaredFieldsInfo(Object instance){
        Map<String,List> map = new HashMap<>();
        List<String> headers = new ArrayList<>();
        List<String> names = new ArrayList<>();
        Assert.notNull(instance,"对象为空");
        Class<?> clazz = instance.getClass();
        Field[] fields=clazz.getDeclaredFields();
        for (int i = 0; i < fields.length ; i++) {
            Field field = fields[i];
            if(field.isAnnotationPresent(FieldName.class)){
                String val = field.getAnnotation(FieldName.class).value();
                Assert.notNull(val,"注解值为空");
                headers.add(val);
                names.add(field.getName());
            }
        }
        map.put(Constants.NAME,names);
        map.put(Constants.HEAD,headers);
        return map;
    }
}
