package com.gitee.poiutils.util;

/**
 * @author Administrator
 * @date 2020/9/9
 * @Description: TODO
 */
public class ExcelCell {

    /**
     * 索引
     */
    private String cellIndex;
    /**
     * 类型：下拉、时间、文本框
     */
    private String type;
    /**
     * 默认值
     */
    private String defaultVal;
    /**
     * 是否必填
     */
    private boolean required;
    /**
     * 类型下拉时的枚举
     */
    private String enums;
    /**
     * 文本框时正则表达式
     * 时间选项时 为时间格式
     */
    private String regularText;
    /**
     * 写入ID预留
     */
    private Integer writeId;
    /**
     * 对比ID 预留
     */
    private Integer readId;
    /**
     * 图片地址
     */
    private String imgPath;
    /**
     * 公式
     */
    private String formula;
    /**
     * 自定义样式
     */
    private String style;

    /**
     * 录入值
     */
    private String cellVal;

    public String getCellVal() {
        //为空优先返回默认值
        return cellVal == null ? defaultVal : cellVal;
    }

    public void setCellVal(String cellVal) {
        this.cellVal = cellVal;
    }

    public String getStyle() {
        return style;
    }

    public void setStyle(String style) {
        this.style = style;
    }

    public String getCellIndex() {
        return cellIndex;
    }

    public void setCellIndex(String cellIndex) {
        this.cellIndex = cellIndex;
    }

    public String getType() {
        return type;
    }

    public void setType(String type) {
        this.type = type;
    }

    public String getDefaultVal() {
        return defaultVal == null ? "" : defaultVal;
    }

    public void setDefaultVal(String defaultVal) {
        this.defaultVal = defaultVal;
    }

    public boolean isRequired() {
        return required;
    }

    public void setRequired(boolean required) {
        this.required = required;
    }

    public String getEnums() {
        return enums;
    }

    public void setEnums(String enums) {
        this.enums = enums;
    }

    public String getRegularText() {
        return regularText;
    }

    public void setRegularText(String regularText) {
        this.regularText = regularText;
    }

    public Integer getWriteId() {
        return writeId;
    }

    public void setWriteId(Integer writeId) {
        this.writeId = writeId;
    }

    public Integer getReadId() {
        return readId;
    }

    public void setReadId(Integer readId) {
        this.readId = readId;
    }

    public String getImgPath() {
        return imgPath;
    }

    public void setImgPath(String imgPath) {
        this.imgPath = imgPath;
    }

    public String getFormula() {
        return formula;
    }

    public void setFormula(String formula) {
        this.formula = formula;
    }
}
