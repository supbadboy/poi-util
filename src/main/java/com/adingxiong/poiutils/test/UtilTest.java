package com.adingxiong.poiutils.test;



import com.adingxiong.poiutils.code.ExcelExport;
import com.adingxiong.poiutils.code.ExcelImport;
import com.adingxiong.poiutils.constant.Constants;
import com.adingxiong.poiutils.util.ExcelUtil;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

/**
 * @ClassName UtilTest
 * @Description TODO
 * @Author xiongchao
 * @Date 2020/11/26 13:50
 **/
public class UtilTest {


    public static void main(String[] args) throws IOException {
        /*commonExport();
        testCostomers();*/
        testImport();
    }

    /**
     *  通用导出测试
     * @throws IOException
     */
    private static void commonExport() throws IOException {
        Workbook wb = new XSSFWorkbook();
        List<ProjectVo> list = mockData();
        ExcelExport.getInstance().fillCommon(wb,list,null);
        FileOutputStream out = new FileOutputStream("C:\\Users\\Administrator\\Desktop\\测试导出.xlsx");
        wb.write(out);
        out.close();
    }
    public static List<ProjectVo> mockData (){
        List<ProjectVo> vo = new ArrayList<>();
        for (int i = 0; i < 15 ; i++) {
            ProjectVo e = new ProjectVo();
            e.setCycle(i + 100 + "");
            Date date = new Date();
            e.setName("waahahah");
            date.setDate(i);
            e.setDate(date);
            e.setMoney(5000.00 + i);
            e.setPerson("张三" + i);
            e.setPhone("15071385455");
            e.setProcessPeople("你是真的批");
            vo.add(e);
        }
        return vo;
    }
    /**
     * 定制导出测试
     * @throws IOException
     */
    public static void testCostomers() throws IOException {
        //定制excel
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("定制表头导出");
        Sheet sheet1 = wb.createSheet("多sheet情景导出");
        CellStyle cellStyle = wb.createCellStyle();
        Font font = wb.createFont();
        List<ProjectVo> list = mockData();

        Map<String, List> declaredFieldsInfo = com.adingxiong.poiutils.util.ClassUtils.getDeclaredFieldsInfo(list.get(0));
        int size = declaredFieldsInfo.get(Constants.HEAD).size();
        sheet.addMergedRegion(new CellRangeAddress(0,0,0,size -1 ));
        //主体部分
        Row row = sheet.createRow(0);
        Cell cell = row.createCell(0);
        font.setBold(true);
        font.setFontHeightInPoints((short) 16);
        font.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);

        CellStyle headStyle = wb.createCellStyle();
        headStyle.setFont(font);
        //设置对齐样式
        headStyle.setAlignment(XSSFCellStyle.ALIGN_CENTER);
        headStyle.setFillForegroundColor(HSSFColor.DARK_RED.index);
        cell.setCellType(1);
        cell.setCellValue("这是定制的单元格头");
        cell.setCellStyle(headStyle);
        ExcelExport.getInstance().fillCustomer(sheet1,0,null,null,list);
        ExcelExport.getInstance().fillCustomer(sheet,1,null,cellStyle,list);
        FileOutputStream out = new FileOutputStream("C:\\Users\\Administrator\\Desktop\\定制导出.xlsx");
        wb.write(out);
        out.close();
    }


    private static void testImport() throws IOException {
        String fileUrl = "C:\\Users\\Administrator\\Desktop\\测试导出.xlsx";
        Workbook workbook = ExcelUtil.readExcel(fileUrl, ProjectVo.class);
        Sheet sheet = workbook.getSheet("sheet");
        List <ProjectVo> list = ExcelImport.getInstance().setRowNum(0).setFieldRows("rows").isFormatTitle(true).setFieldError("error").transformation(sheet ,ProjectVo.class);
        list.stream().forEach(e ->{
            System.out.println(e.toString());
        });

    }


}
