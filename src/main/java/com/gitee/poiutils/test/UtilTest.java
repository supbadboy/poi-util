package com.gitee.poiutils.test;



import com.gitee.poiutils.code.ExcelExport;
import com.gitee.poiutils.code.ExcelImport;
import com.gitee.poiutils.code.WordExport;
import com.gitee.poiutils.constant.Constants;
import com.gitee.poiutils.util.ExcelUtil;
import com.gitee.poiutils.util.ClassUtils;
import lombok.Builder;
import lombok.Data;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.file.Path;
import java.util.*;

/**
 * ClassName UtilTest
 * Description TODO
 * @author xiongchao
 * Date 2020/11/26 13:50
 **/
public class UtilTest {


    public static void main(String[] args) throws IOException {
        /*commonExport();
        testCostomers();*/
        //testImport();
        testWordExport();
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

        Map<String, List> declaredFieldsInfo = ClassUtils.getDeclaredFieldsInfo(list.get(0));
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
        List <ProjectVo> list = ExcelImport.getInstance()
                .setRowNum(0)
                .setFieldRows("rows")
                .isFormatTitle(true)
                .setFieldError("error")
                .transformation(sheet ,ProjectVo.class);
        list.stream().forEach(e ->{
            System.out.println(e.toString());
        });

    }


    private static void testWordExport() {
        File template = new File("C:\\Users\\Administrator\\Desktop\\repair_report.docx");
        InputStream btImg = ExcelUtil.getUrl("http://192.168.2.126/group1/M00/45/79/wKgCfl8PrUyAMNXUAAErNrq-Q9c236.jpg");
        Map<String,Object> resMap = new HashMap<>();
        resMap.put("username" , "测试word模版导出");
        resMap.put("nickname", "昵称");
        resMap.put("date", Constants.simpleDateFormat.format(new Date()));
        resMap.put("email" ,"1374543195@qq.com");
        resMap.put("note","七大爷八大姑子稀里糊涂叽里呱啦");
        try {
            resMap.put("img" ,btImg);
            FileInputStream fileInputStream = new FileInputStream(template);
            List<UserInfo> list = new ArrayList<>();
            list.add(UserInfo.builder()
                    .name("王中华")
                    .age("25")
                    .phone("15071384121")
                    .email("1119624186@qq.com")
                    .job("高级软件开发工程师")
                    .build()
            );
            list.add(UserInfo.builder()
                    .name("四匹马")
                    .age("25")
                    .phone("15071384121")
                    .email("1119624186@qq.com")
                    .job("ui不错哦")
                    .build());
            resMap.put("list",list);

            Path path = WordExport.getInstance().convertWord(fileInputStream ,resMap ,"模版文件到处.docx" );
            System.out.println(path);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {

        }
    }
    @Data
    @Builder
     private static class UserInfo {
        private String name;
        private String age ;
        private String phone;
        private String email;
        private String job;


    }
}
