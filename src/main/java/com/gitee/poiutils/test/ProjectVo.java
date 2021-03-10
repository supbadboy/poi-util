package com.gitee.poiutils.test;

import com.gitee.poiutils.interfaces.FieldName;
import lombok.Data;

import java.util.Date;

/**
 * ClassName ProjectVo
 * Description TODO
 * @author xiongchao
 * Date 2020/11/26 15:48
 **/
@Data
public class ProjectVo {



    @FieldName(value = "项目")
    private String name;

    @FieldName(value = "电话" ,pattern = "/^(13[0-9]|14[01456879]|15[0-35-9]|16[2567]|17[0-8]|18[0-9]|19[0-35-9])\\d{8}$/")
    private String phone;

    @FieldName(value = "联系人")
    private String person;

    @FieldName(value = "金额")
    private Double money;

    @FieldName(value = "负责人" ,required = true)
    private String processPeople;

    @FieldName(value = "周期")
    private String cycle;

    @FieldName(value = "记录日期",dateFormat = "yyyy-MM-dd")
    private Date date;

    private String error;

    private Integer rows;
}
