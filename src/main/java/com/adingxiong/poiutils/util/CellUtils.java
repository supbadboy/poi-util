package com.adingxiong.poiutils.util;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;

import java.util.LinkedHashMap;
import java.util.Map;

/**
 * @ClassName CellUtils
 * @Description  单元格操作的工具类
 *
 *         TODO  : 需要优化的地方 这里需要考虑的是  在创建单元格的时候,设置每个单元格的样式,存在多次创建CellStyle 对象
 *          和 Font 对象  ,后期要根据数据量来考虑是否将创建对象改为clone (复杂的对象 clone效率更高,但是性能还没分析
 * @Author xiongchao
 * @Date 2020/11/27 9:46
 **/
public class CellUtils {


    /**
     * 创建单元格表头样式
     *
     * @param workbook 工作薄
     */
    public static CellStyle createCellHeadStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        // 设置边框样式
        style.setBorderBottom(XSSFCellStyle.BORDER_THIN);
        style.setBorderLeft(XSSFCellStyle.BORDER_THIN);
        style.setBorderRight(XSSFCellStyle.BORDER_THIN);
        style.setBorderTop(XSSFCellStyle.BORDER_THIN);
        //设置对齐样式
        style.setAlignment(XSSFCellStyle.ALIGN_CENTER);
        // 生成字体
        Font font = workbook.createFont();
        // 表头样式
        style.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
        font.setFontHeightInPoints((short) 16);
        font.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
        // 把字体应用到当前的样式
        style.setFont(font);
        return style;
    }

    /**
     * 创建单元格标题样式
     * @param workbook
     * @return
     */
    public static CellStyle createCellTitleStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        // 设置边框样式
        style.setBorderBottom(XSSFCellStyle.BORDER_THIN);
        style.setBorderLeft(XSSFCellStyle.BORDER_THIN);
        style.setBorderRight(XSSFCellStyle.BORDER_THIN);
        style.setBorderTop(XSSFCellStyle.BORDER_THIN);
        //设置对齐样式
        style.setAlignment(XSSFCellStyle.ALIGN_CENTER);
        // 生成字体
        Font font = workbook.createFont();
        // 表头样式
        style.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
        style.setFillForegroundColor(HSSFColor.DARK_BLUE.index);
        font.setFontHeightInPoints((short) 12);
        font.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
        font.setColor(HSSFColor.WHITE.index);
        // 把字体应用到当前的样式
        style.setFont(font);
        return style;
    }


    /**
     * 创建单元格正文样式
     *
     * @param workbook 工作薄
     */
    public static CellStyle createCellContentStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        // 设置边框样式
        style.setBorderBottom(XSSFCellStyle.BORDER_THIN);
        style.setBorderLeft(XSSFCellStyle.BORDER_THIN);
        style.setBorderRight(XSSFCellStyle.BORDER_THIN);
        style.setBorderTop(XSSFCellStyle.BORDER_THIN);
        //设置对齐样式
        style.setAlignment(XSSFCellStyle.ALIGN_CENTER);
        //自动换行
        style.setWrapText(true);
        // 生成字体
        Font font = workbook.createFont();
        // 正文样式
        style.setFillPattern(XSSFCellStyle.NO_FILL);
        style.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
        font.setBoldweight(XSSFFont.BOLDWEIGHT_NORMAL);
        // 把字体应用到当前的样式
        style.setFont(font);
        return style;
    }


    /**
     * 单元格样式列表
     */
    public static Map< String, CellStyle> styleMap(Workbook workbook) {
        Map < String, CellStyle > styleMap = new LinkedHashMap<String, CellStyle >();
        styleMap.put("head", createCellHeadStyle(workbook));
        styleMap.put("title", createCellTitleStyle(workbook));
        styleMap.put("content", createCellContentStyle(workbook));
        return styleMap;
    }
}
