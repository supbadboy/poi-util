package com.gitee.poiutils.code;

import com.gitee.poiutils.constant.Constants;
import com.gitee.poiutils.constant.Errorcons;
import com.gitee.poiutils.util.CellUtils;
import com.gitee.poiutils.util.ClassUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.springframework.util.CollectionUtils;
import org.springframework.util.StringUtils;

import java.lang.reflect.Field;
import java.text.SimpleDateFormat;
import java.time.Instant;
import java.util.*;

/**
 * ClassName ExcelExport
 * Description excel 通用导出模块
 * @author xiongchao
 * Date 2020/11/26 11:04
 **/
public class ExcelExport {

    /**
     *  ThreadLocal 线程内部特定的存储类  不同线程的变量是互相隔离的  每个线程可以操作属于自己的那个 变量值
     *  可以进行get  set操作 并且 不会影响到其它线程对变量的使用
     */
    private ThreadLocal<Integer> indexRow = new ThreadLocal<>();

    private ThreadLocal<SimpleDateFormat> sdf = new ThreadLocal<>();

    private static ExcelExport excelExport;

    private final Map<Integer,Integer> cellMaxWidth = new HashMap<>();

    public ExcelExport setIndex(Integer index){
        this.indexRow.set(index);
        return this;
    }

    public ExcelExport setSimpleDateFormat (SimpleDateFormat sdf){
        this.sdf.set(sdf);
        return this;
    }

    public static synchronized ExcelExport getInstance (){
        if(excelExport  == null) {
            excelExport = new ExcelExport();
        }
        return excelExport;
    }

    /**
     *  通用导入   默认标题字体加粗 字体居中  带有单元格边框线
     * @param wb  wb对象
     * @param list 需要填充的集合数据
     * @param sheetName   sheet名称
     * @param <T> 泛型
     */
    public synchronized <T>  void fillCommon (Workbook wb , List<T>list ,String sheetName){
        if(wb == null) {
            throw new NullPointerException(Errorcons.WK_EMPTY);
        }
        if (StringUtils.isEmpty(sheetName)) {
            sheetName = Constants.DF_SHEET;
        }
        Sheet sheet = wb.createSheet(sheetName);
        CellStyle cellStyle = wb.createCellStyle();
        //手动填充索引标题列  调用地方无需进行任何处理
        if(CollectionUtils.isEmpty(list)){
            throw new NullPointerException(Errorcons.COLL_EMPTY);
        }
        Map<String, List> fieldsInfo = ClassUtils.getDeclaredFieldsInfo(list.get(0));
        List indexList = fieldsInfo.get(Constants.NAME);
        List<String> headers = fieldsInfo.get(Constants.HEAD);

        if(CollectionUtils.isEmpty(indexList) || CollectionUtils.isEmpty(headers)){
            throw new ClassCastException(Errorcons.ANAY_ERROR);
        }
        try {
            SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
            Integer indexRow = Integer.valueOf(0);
            if (this.indexRow.get()  != null) {
                sdf = this.sdf.get();
                indexRow = this.indexRow.get();
            }
            Row row = sheet.getRow(indexRow);
            Cell cell = null;

            if(row == null) {
                //插入标题
                row = sheet.createRow(indexRow);
                for (int i = 0; i < headers.size(); i++) {
                    createCell(row ,i ,headers.get(i),CellUtils.createCellTitleStyle(wb));
                }
            }
            //跳过标题
            for (int i = 0; i < list.size(); i++) {
                Object t = list.get(i);
                row = sheet.getRow(i + indexRow.intValue() + 1 );
                if(row  == null) {
                    row = sheet.createRow(i + indexRow.intValue() + 1);
                }
                for (int j = 0; j < indexList.size(); j++) {
                    if(!StringUtils.isEmpty(indexList.get(j))){
                        Field filed = t.getClass().getDeclaredField(indexList.get(j).toString());
                        filed.setAccessible(true);
                        String cellValue = "";
                        String fileType = filed.getType().getSimpleName();
                        cellValue = getString(sdf, t, filed, cellValue, fileType);
                        createCell(row, j ,cellValue,CellUtils.createCellContentStyle(wb));
                    }
                }
            }
            //设置单元列宽度  为最大宽度
            for (int i = 0; i < headers.size(); i++) {
                sheet.setColumnWidth(i , cellMaxWidth.get(i));
            }
        } catch (NoSuchFieldException e) {
            e.printStackTrace();
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        }  finally {
            this.indexRow.remove();
            this.sdf.remove();
        }
    }



    /**
     *  创建单元格
     * @param row  行
     * @param index 索引
     * @param value 填充的值
     */
    private void createCell(Row row ,Integer index ,String value ,CellStyle style){
        Cell cell = row.getCell(index);
        if(cell == null) {
            cell = row.createCell(index);
        }
        cell.setCellType(1);
        cell.setCellValue(value);
        if(style != null) {
            cell.setCellStyle(style);
        }

        if(!StringUtils.isEmpty(value)){
            int currentSize = value.length()*2*256;
            if(cellMaxWidth.isEmpty()){
                cellMaxWidth.put(index,currentSize);
            }else {
                Integer integer = cellMaxWidth.get(index);
                if(integer == null){
                    cellMaxWidth.put(index,currentSize);
                }else{
                    if(integer < currentSize) {
                        cellMaxWidth.put(index , currentSize);
                    }
                }
            }
        }
    }


    /**
     *  定制导出excel 模块 , 可以自定义excel表头,自定义sheet页数 (通常用于多个sheet页面的场景) 字体和样式可以不填
     * @param sheet   sheet 对象
     * @param skipRow  跳过行数
     * @param font  字体对象
     * @param cellStyle   单元格样式
     * @param list  需要导入的数据集合
     * @param <T> 泛型
     */
    public synchronized <T> void fillCustomer(Sheet sheet , int skipRow , Font font , CellStyle cellStyle , List<?> list){
        if(CollectionUtils.isEmpty(list)){
            throw new NullPointerException(Errorcons.COLL_EMPTY);
        }

        if(cellStyle != null) {
            cellStyle.setBorderBottom(XSSFCellStyle.BORDER_THIN);
            cellStyle.setBorderLeft(XSSFCellStyle.BORDER_THIN);
            cellStyle.setBorderRight(XSSFCellStyle.BORDER_THIN);
            cellStyle.setBorderTop(XSSFCellStyle.BORDER_THIN);
            //设置对齐样式
            cellStyle.setAlignment(XSSFCellStyle.ALIGN_CENTER);
            cellStyle.setWrapText(true);
        }
        Map<String, List> fieldsInfo = ClassUtils.getDeclaredFieldsInfo(list.get(0));
        List indexList = fieldsInfo.get(Constants.NAME);
        List<String> headers = fieldsInfo.get(Constants.HEAD);

        if(CollectionUtils.isEmpty(indexList) || CollectionUtils.isEmpty(headers)){
            throw new ClassCastException(Errorcons.ANAY_ERROR);
        }
        try {
            SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
            Integer indexRow = Integer.valueOf(0);
            if (this.indexRow.get()  != null) {
                sdf = this.sdf.get();
                indexRow = this.indexRow.get();
            }
            Row row = sheet.getRow(indexRow  + skipRow);
            Cell cell = null;

            if(row == null) {
                //插入标题
                row = sheet.createRow(indexRow + skipRow);
                for (int i = 0; i < headers.size(); i++) {
                    createCell(row ,i ,headers.get(i),cellStyle);
                }
            }
            //跳过标题
            for (int i = 0; i < list.size(); i++) {
                Object t = list.get(i);
                row = sheet.getRow(i + indexRow.intValue() + 1 + skipRow );
                if(row  == null) {
                    row = sheet.createRow(i + indexRow.intValue() + 1 + skipRow);
                }
                for (int j = 0; j < indexList.size(); j++) {
                    if(!StringUtils.isEmpty(indexList.get(j))){
                        Field filed = t.getClass().getDeclaredField(indexList.get(j).toString());
                        filed.setAccessible(true);
                        String cellValue = "";
                        String fileType = filed.getType().getSimpleName();
                        cellValue = getString(sdf, t, filed, cellValue, fileType);
                        createCell(row, j ,cellValue,cellStyle);
                    }
                }
            }
            //设置单元列宽度  为最大宽度
            for (int i = 0; i < headers.size(); i++) {
                sheet.setColumnWidth(i , cellMaxWidth.get(i));
            }
        } catch (NoSuchFieldException e) {
            e.printStackTrace();
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        }  finally {
            this.indexRow.remove();
            this.sdf.remove();
        }
    }

    private String getString(SimpleDateFormat sdf, Object t, Field filed, String cellValue, String fileType) throws IllegalAccessException {
        switch (fileType){
            case "Date":
                Date date = (Date) filed.get(t);
                if(date != null) {
                    cellValue = sdf.format(date);
                }
                break;
            case "Instant":
                Instant instant = (Instant) filed.get(t);
                if(instant != null) {
                    cellValue = sdf.format(Date.from(instant));
                }
                break;
            default:
                Object obj = filed.get(t);
                if(obj != null) {
                    cellValue = obj.toString();
                }
                break;
        }
        return cellValue;
    }


}
