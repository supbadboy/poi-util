package com.gitee.poiutils.util;

import com.gitee.poiutils.interfaces.FieldName;
import com.gitee.poiutils.constant.Constants;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.util.Assert;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.net.HttpURLConnection;
import java.net.URL;
import java.text.NumberFormat;
import java.util.HashMap;
import java.util.Map;

/**
 * ClassName ExcelUtil
 * Description TODO
 * @author xiongchao
 * Date 2020/11/27 14:46
 **/
public class ExcelUtil {
    public static final NumberFormat numberFormat = NumberFormat.getInstance();

    static  {
        numberFormat.setMaximumFractionDigits(2);
        numberFormat.setGroupingUsed(true);
    }

    public static Map<String, Integer> rowMap(Row row, Field[] fields, Boolean formatTitle)
    {
        Map resultMap = new HashMap();
        Map fieldMap = new HashMap();
        int cellNum = row.getPhysicalNumberOfCells();
        for (Field field : fields) {
            if (field.isAnnotationPresent(FieldName.class)) {
                FieldName annotation = (FieldName)field.getAnnotation(FieldName.class);
                if (formatTitle.booleanValue()) {
                    fieldMap.put(annotation.value().trim().replaceAll("\r|\n", ""), field.getName());
                } else {
                    fieldMap.put(annotation.value(), field.getName());
                }
            }
        }
        for (int i = 0; i < cellNum; i++) {
            Cell cell = row.getCell(i);
            if (cell != null) {
                cell.setCellType(1);
                String cellValue = getCellValue(cell);
                if (formatTitle.booleanValue()) {
                    cellValue = getCellValue(cell).trim().replaceAll("\r|\n", "");
                }
                if (fieldMap.containsKey(cellValue)) {
                    resultMap.put(fieldMap.get(cellValue), Integer.valueOf(i));
                }
            }
        }
        return resultMap;
    }

    public static String getCellValue(Cell cell)
    {
        String value = "";
        switch (cell.getCellType()) {
            case 0:
                if (DateUtil.isCellDateFormatted(cell)) {
                    value = Constants.simpleDateFormat.format(HSSFDateUtil.getJavaDate(cell.getNumericCellValue()));
                }else {
                    value = numberFormat.format(cell.getNumericCellValue());
                }
                break;
            case 1:
                value = cell.getStringCellValue();
                break;
            case 4:
                value = cell.getBooleanCellValue() + "";
                break;
            case 2:
                value = cell.getNumericCellValue() + "";
                break;
            case 3:
                value = "";
                break;
        }

        return value;
    }

    public static Workbook readExcel(InputStream inputStream, String fileName) throws IOException {
        Assert.notNull(fileName, "文件后缀名不能为空");
        Assert.notNull(inputStream, "输入流不能为空");

        String extensionName = "xlsx";
        Workbook workbook;
        if (extensionName.endsWith(fileName)) {
            workbook = new XSSFWorkbook(inputStream);
        } else {
            workbook = new HSSFWorkbook(inputStream);
        }
        return workbook;
    }

    public static Workbook readExcel(String filePath, Class clazz) throws IOException {
        Workbook wb = null;
        Assert.notNull(filePath, "链接地址不能为空");
        String extString = filePath.substring(filePath.lastIndexOf('.'));
        InputStream is = null;
        try {
            is = clazz.getResourceAsStream(filePath);
            if (is == null) {
                is = new FileInputStream(filePath);
            }
            String extensionName = ".xlsx";
            if (extensionName.endsWith(extString)) {
                wb = new XSSFWorkbook(is);
            }else {
                wb = new HSSFWorkbook(is);
            }
        }
        catch (FileNotFoundException e) {
            throw new IOException("File does not exist");
        } catch (IOException e) {
            throw new IOException("Resolve excel error");
        }
        is.close();
        return wb;
    }

    public static InputStream getUrl(String imgUrl) {
        InputStream inStream = null;
        try {
            URL url = new URL(imgUrl);
            HttpURLConnection conn = (HttpURLConnection)url.openConnection();
            conn.setRequestMethod("GET");
            conn.setConnectTimeout(5000);
            inStream = conn.getInputStream();
        }
        catch (Exception localException) {
        }
        return inStream;
    }
}
