package com.gitee.poiutils.code;

import com.alibaba.fastjson.JSON;
import com.gitee.poiutils.handler.DefHandler;
import com.gitee.poiutils.handler.FormulaHandler;
import com.gitee.poiutils.handler.InputHandler;
import com.gitee.poiutils.handler.SelectHandler;
import com.gitee.poiutils.util.ExcelCell;
import com.gitee.poiutils.util.ExcelTools;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.util.StringUtils;

import java.io.*;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.ConcurrentMap;

/**
 * @ClassName ExcelToHtml
 * @Description  excel  转html
 * @Author xiongchao
 * @Date 2020/10/21 13:19
 **/
public class ExcelToHtml {

    public static ExcelToHtml getInstance(){
        return new ExcelToHtml();
    }

    private static ConcurrentMap<String,DefHandler> handlers = new ConcurrentHashMap<>();

    static String[] bordesr={"border-top:","border-right:","border-bottom:","border-left:"};

    static String[] borderStyles={"solid ","solid ","solid ","solid ","solid ","solid ","solid ","solid ","solid ","solid","solid","solid","solid","solid"};


    /**
     *  Excel 资源转H5 页面
     * @param path  目标资源路径
     * @param wb wb对象
     * @param isWithStyle 是否保留源文件样式
     * @param rowSize 固定行数
     * @return
     */
    public String excelToH5(String path ,Workbook wb ,boolean isWithStyle,int rowSize){
        handlers.put("input",new InputHandler());
        handlers.put("select",new SelectHandler());
        handlers.put("formula",new FormulaHandler());
        Map<String,String> param = new HashMap<>();
        if( !StringUtils.isEmpty(path)){
            return readExcelToH5(path,param,isWithStyle,rowSize);
        }
        if( wb != null) {
            return readExcelToH5(wb,param,isWithStyle,rowSize);
        }
        return "未找到目标文件";
    }

    public String readExcelToH5(String path,Map<String,String> param,boolean isWithStyle,int rowSize) {
        InputStream is = null;
        String htmlExcel = null;
        try {
            File sourcefile = new File(path);
            is = new FileInputStream(sourcefile);
            Workbook wb = WorkbookFactory.create(is);
            htmlExcel = getString(wb, param, isWithStyle, htmlExcel,rowSize);
        } catch (Exception e) {
            e.printStackTrace();
        }finally{
            try {
                if(is!=null){
                    is.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return htmlExcel;
    }


    /**
     * 程序入口方法
     * @param
     * @param isWithStyle 是否需要表格样式 包含 字体 颜色 边框 对齐方式
     * @return <table>...</table> 字符串
     */
    public  String readExcelToH5(Workbook wb , Map<String,String> param, boolean isWithStyle,int rowSize){
        String htmlExcel = null;
        try {
            htmlExcel = getString(wb, param, isWithStyle, htmlExcel,rowSize);
        } catch (Exception e) {
            e.printStackTrace();
        }
        return htmlExcel;
    }

    private  String getString(Workbook wb, Map<String, String> param, boolean isWithStyle, String htmlExcel,int rowSize) {
        if (wb instanceof XSSFWorkbook) {
            XSSFWorkbook xWb = (XSSFWorkbook) wb;
            htmlExcel = ExcelToHtml.getExcelInfo(xWb,param,isWithStyle,rowSize);
        }else if(wb instanceof HSSFWorkbook){
            HSSFWorkbook hWb = (HSSFWorkbook) wb;
            htmlExcel = ExcelToHtml.getExcelInfo(hWb,param,isWithStyle,rowSize);
        }
        return htmlExcel;
    }

    /**
     * 拼接HTML
     * @param wb
     * @param isWithStyle
     * @return
     */
    private static String getExcelInfo(Workbook wb, Map<String,String> param, boolean isWithStyle , int rowSize){
        if(rowSize == 0) {
            rowSize = 1;
        }
        StringBuffer sb = new StringBuffer();
        //获取第一个Sheet的内容
        Sheet sheet = wb.getSheetAt(0);
        int lastRowNum = sheet.getLastRowNum();
        Map<String, String> map[] = getRowSpanColSpanMap(sheet);
        //计算表格宽度 找最大的一条作为table 宽度 避免出现横向拉伸  这里去掉 宽度写死
        sb.append("<table style='border-collapse:collapse;' width='98.8%'>");
        //兼容
        Row row = null;
        Cell cell = null;
        //前2行固定为表头
        for (int rowNum = sheet.getFirstRowNum(); rowNum <= lastRowNum; rowNum++) {
            row = sheet.getRow(rowNum);
            //这种情况不考虑了
            if (row == null) {
                sb.append("<tr><td > &nbsp;</td></tr>");
                continue;
            }
            if(rowNum == rowSize) {
                sb.append("</table><div style='width:100%;height:85%;overflow-x:hidden; overflow-y:scroll;'><table style='border-collapse:collapse;' width='100%'>");
            }

            //设置行高度
            float rowHeight = row.getHeightInPoints();
            sb.append("<tr style='height:"+rowHeight+"px;'>");
            int lastColNum = row.getLastCellNum();
            for (int colNum = 0; colNum < lastColNum; colNum++) {
                cell = row.getCell(colNum);
                String key = ExcelTools.excelColIndexToStr(colNum,rowNum);
                String def = "&nbsp;";
                //获取批注
                if(cell != null && cell.getCellComment() != null){
                    String cellComment = cell.getCellComment().getString().toString();
                    ExcelCell excelCell = JSON.parseObject(cellComment, ExcelCell.class);
                    if(param != null && param.containsKey(key)){
                        excelCell.setCellVal(param.get(key));
                    }
                    def = DefHandler.handlers.get(excelCell.getType()).excute(excelCell,param);
                }
                //特殊情况 空白的单元格会返回null
                if (cell == null) {
                    sb.append("<td id=" + key + "  title=" + key +" >" + def + "</td>");
                    continue;
                }
                String stringValue = getCellValue(cell);
                if (map[0].containsKey(rowNum + "," + colNum)) {
                    String pointString = map[0].get(rowNum + "," + colNum);
                    map[0].remove(rowNum + "," + colNum);
                    int bottomeRow = Integer.valueOf(pointString.split(",")[0]);
                    int bottomeCol = Integer.valueOf(pointString.split(",")[1]);
                    int rowSpan = bottomeRow - rowNum + 1;
                    int colSpan = bottomeCol - colNum + 1;
                    sb.append("<td rowspan= '" + rowSpan + "' colspan= '"+ colSpan + "' ");
                } else if (map[1].containsKey(rowNum + "," + colNum)) {
                    map[1].remove(rowNum + "-" + colNum);
                    continue;
                } else {
                    sb.append("<td ");
                }

                //判断是否需要样式
                if(isWithStyle){
                    //处理单元格样式
                    dealExcelStyle(wb, sheet, cell, sb);
                }
                sb.append(" id=" + key +" ");
                sb.append(" title=" + key +" ");
                sb.append(">");
                if (stringValue == null || "".equals(stringValue.trim())) {
                    sb.append(def);
                } else {
                    // 将ascii码为160的空格转换为html下的空格（&nbsp;）
                    if(param != null && param.containsKey(key)){
                        sb.append(def);
                    } else {
                        sb.append(stringValue.replace(String.valueOf((char) 160),"&nbsp;"));
                    }
                }
                sb.append("</td>");
            }
            sb.append("</tr>");
        }
        sb.append("</div></table>");
        return sb.toString();
    }

    /**
     * 合并单元格
     * @param sheet
     * @return
     */
    private static Map<String, String>[] getRowSpanColSpanMap(Sheet sheet) {
        Map<String, String> map0 = new HashMap<String, String>();
        Map<String, String> map1 = new HashMap<String, String>();
        int mergedNum = sheet.getNumMergedRegions();
        CellRangeAddress range = null;
        for (int i = 0; i < mergedNum; i++) {
            range = sheet.getMergedRegion(i);
            int topRow = range.getFirstRow();
            int topCol = range.getFirstColumn();
            int bottomRow = range.getLastRow();
            int bottomCol = range.getLastColumn();
            map0.put(topRow + "," + topCol, bottomRow + "," + bottomCol);
            int tempRow = topRow;
            while (tempRow <= bottomRow) {
                int tempCol = topCol;
                while (tempCol <= bottomCol) {
                    map1.put(tempRow + "," + tempCol, "");
                    tempCol++;
                }
                tempRow++;
            }
            map1.remove(topRow + "," + topCol);
        }
        Map[] map = { map0, map1 };
        return map;
    }


    /**
     * 获取表格单元格Cell内容
     * @param cell
     * @return
     */
    private static String getCellValue(Cell cell) {
        String result = new String();
        switch (cell.getCellType()) {
            // 数字类型
            case Cell.CELL_TYPE_NUMERIC:
                // 处理日期格式、时间格式
                if (HSSFDateUtil.isCellDateFormatted(cell)) {
                    SimpleDateFormat sdf = null;
                    if (cell.getCellStyle().getDataFormat() == HSSFDataFormat.getBuiltinFormat("h:mm")) {
                        sdf = new SimpleDateFormat("HH:mm");
                    } else {
                        sdf = new SimpleDateFormat("yyyy-MM-dd");
                    }
                    Date date = cell.getDateCellValue();
                    result = sdf.format(date);
                } else if (cell.getCellStyle().getDataFormat() == 58) {
                    // 处理自定义日期格式：m月d日(通过判断单元格的格式id解决，id的值是58)
                    SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
                    double value = cell.getNumericCellValue();
                    Date date = DateUtil
                        .getJavaDate(value);
                    result = sdf.format(date);
                } else {
                    double value = cell.getNumericCellValue();
                    CellStyle style = cell.getCellStyle();
                    DecimalFormat format = new DecimalFormat();
                    String temp = style.getDataFormatString();
                    // 单元格设置成常规
                    if (temp.equals("General")) {
                        format.applyPattern("#");
                    }
                    result = format.format(value);
                }
                break;
            // String类型
            case Cell.CELL_TYPE_STRING:
                result = cell.getRichStringCellValue().toString().replace("\n","<br>");
                break;
            case Cell.CELL_TYPE_BLANK:
                result = "";
                break;
            default:
                result = "";
                break;
        }
        return result;
    }

    /**
     * 处理表格样式
     * @param wb
     * @param sheet
     * @param cell
     * @param sb
     */
    private static void dealExcelStyle(Workbook wb, Sheet sheet, Cell cell, StringBuffer sb){
        CellStyle cellStyle = cell.getCellStyle();
        if (cellStyle != null) {
            short alignment = cellStyle.getAlignment();
            sb.append("align='" + convertAlignToHtml(alignment) + "' ");//单元格内容的水平对齐方式
            short verticalAlignment = cellStyle.getVerticalAlignment();
            sb.append("valign='"+ convertVerticalAlignToHtml(verticalAlignment)+ "' ");//单元格中内容的垂直排列方式
            if (wb instanceof XSSFWorkbook) {
                XSSFFont xf = ((XSSFCellStyle) cellStyle).getFont();
                short boldWeight = xf.getBoldweight();
                sb.append("style='");
                sb.append("font-weight:" + boldWeight + ";"); // 字体加粗
                sb.append("font-size: " + xf.getFontHeightInPoints()  + "px;"); // 字体大小
                float columnWidth = sheet.getColumnWidthInPixels(cell.getColumnIndex()) ;
                sb.append("width:" + columnWidth/2+ "px;");
                XSSFColor xc = xf.getXSSFColor();
                if (xc != null && !"".equals(xc)) {
                    sb.append("color:#" + xc.getARGBHex().substring(2) + ";");  // 字体颜色
                }
                XSSFColor bgColor = (XSSFColor) cellStyle.getFillForegroundColorColor();
                if (bgColor != null && !"".equals(bgColor)) {
                    sb.append("background-color:#" + bgColor.getARGBHex().substring(2) + ";"); //
                }
                sb.append(getBorderStyle(0,cellStyle.getBorderTop(), ((XSSFCellStyle) cellStyle).getTopBorderXSSFColor()));
                sb.append(getBorderStyle(1,cellStyle.getBorderRight(), ((XSSFCellStyle) cellStyle).getRightBorderXSSFColor()));
                sb.append(getBorderStyle(2,cellStyle.getBorderBottom(), ((XSSFCellStyle) cellStyle).getBottomBorderXSSFColor()));
                sb.append(getBorderStyle(3,cellStyle.getBorderLeft(), ((XSSFCellStyle) cellStyle).getLeftBorderXSSFColor()));

            } else if(wb instanceof HSSFWorkbook){
                HSSFFont hf = ((HSSFCellStyle) cellStyle).getFont(wb);
                short boldWeight = hf.getBoldweight();
                short fontColor = hf.getColor();
                sb.append("style='");
                HSSFPalette palette = ((HSSFWorkbook) wb).getCustomPalette(); // 绫籋SSFPalette鐢ㄤ簬姹傜殑棰滆壊鐨勫浗闄呮爣鍑嗗舰寮�
                HSSFColor hc = palette.getColor(fontColor);
                sb.append("font-weight:" + boldWeight + ";"); // 瀛椾綋鍔犵矖
                sb.append("font-size: " + hf.getFontHeightInPoints() + "px;"); // 瀛椾綋澶у皬
                String fontColorStr = convertToStardColor(hc);
                if (fontColorStr != null && !"".equals(fontColorStr.trim())) {
                    sb.append("color:" + fontColorStr + ";"); // 瀛椾綋棰滆壊
                }
                float columnWidth = sheet.getColumnWidthInPixels(cell.getColumnIndex()) ;
                sb.append("width:" + columnWidth/2 + "px;");
                short bgColor = cellStyle.getFillForegroundColor();
                hc = palette.getColor(bgColor);
                String bgColorStr = convertToStardColor(hc);
                if (bgColorStr != null && !"".equals(bgColorStr.trim())) {
                    sb.append("background-color:" + bgColorStr + ";"); // 鑳屾櫙棰滆壊
                }
                sb.append( getBorderStyle(palette,0,cellStyle.getBorderTop(),cellStyle.getTopBorderColor()));
                sb.append( getBorderStyle(palette,1,cellStyle.getBorderRight(),cellStyle.getRightBorderColor()));
                sb.append( getBorderStyle(palette,3,cellStyle.getBorderLeft(),cellStyle.getLeftBorderColor()));
                sb.append( getBorderStyle(palette,2,cellStyle.getBorderBottom(),cellStyle.getBottomBorderColor()));
            }

            sb.append("' ");
        }
    }

    /**
     * 单元格内容的水平对齐方式
     * @param alignment
     * @return
     */
    private static String convertAlignToHtml(short alignment) {
        String align = "left";
        switch (alignment) {
            case CellStyle.ALIGN_LEFT:
                align = "left";
                break;
            case CellStyle.ALIGN_CENTER:
                align = "center";
                break;
            case CellStyle.ALIGN_RIGHT:
                align = "right";
                break;
            default:
                break;
        }
        return align;
    }

    /**
     * 单元格中内容的垂直排列方式
     * @param verticalAlignment
     * @return
     */
    private static String convertVerticalAlignToHtml(short verticalAlignment) {
        String valign = "middle";
        switch (verticalAlignment) {
            case CellStyle.VERTICAL_BOTTOM:
                valign = "bottom";
                break;
            case CellStyle.VERTICAL_CENTER:
                valign = "center";
                break;
            case CellStyle.VERTICAL_TOP:
                valign = "top";
                break;
            default:
                break;
        }
        return valign;
    }

    /**
     * 获取颜色
     * @param hc
     * @return
     */
    private static String convertToStardColor(HSSFColor hc) {
        StringBuffer sb = new StringBuffer("");
        if (hc != null) {
            if (HSSFColor.AUTOMATIC.index == hc.getIndex()) {
                return null;
            }
            sb.append("#");
            for (int i = 0; i < hc.getTriplet().length; i++) {
                sb.append(fillWithZero(Integer.toHexString(hc.getTriplet()[i])));
            }
        }
        return sb.toString();
    }

    private static String fillWithZero(String str) {
        if (str != null && str.length() < 2) {
            return "0" + str;
        }
        return str;
    }


    /**
     * 边框
     * @param palette
     * @param b
     * @param s
     * @param t
     * @return
     */
    private static  String getBorderStyle(HSSFPalette palette , int b, short s, short t){
        if(s==0)return  bordesr[b]+borderStyles[s]+"#d0d7e5 1px;";;
        String borderColorStr = convertToStardColor( palette.getColor(t));
        borderColorStr=borderColorStr==null|| borderColorStr.length()<1?"#000000":borderColorStr;
        return bordesr[b]+borderStyles[s]+borderColorStr+" 1px;";

    }

    /**
     * 边框样式
     * @param b
     * @param s
     * @param xc
     * @return
     */
    private static  String getBorderStyle(int b,short s, XSSFColor xc){
        if(s==0) {
            return  bordesr[b]+borderStyles[s]+"#d0d7e5 1px;";
        }
        if (xc != null && !"".equals(xc)) {
            //t.getARGBHex();
            String borderColorStr = xc.getARGBHex();
            borderColorStr=borderColorStr==null|| borderColorStr.length()<1?"#000000":borderColorStr.substring(2);
            return bordesr[b]+borderStyles[s]+borderColorStr+" 1px;";
        }
        return "";
    }

}
