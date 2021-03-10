package com.gitee.poiutils.interfaces;

import java.lang.annotation.*;

/**
 * ClassName CellFiled
 * Description   自定义单元格注解
 * @author xiongchao
 * Date 2020/11/26 13:14
 **/
@Target({java.lang.annotation.ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface FieldName {

    String value() default "";

    String dateFormat() default "";

    boolean required() default false;
    /*
     正则限制
     */
    String pattern() default "";
}
