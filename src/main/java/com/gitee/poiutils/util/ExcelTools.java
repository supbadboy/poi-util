package com.gitee.poiutils.util;


import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;

/**
 * @ClassName ExcelTools
 * @Description  excel  工具类
 * @Author xiongchao
 * @Date 2020/11/9 9:43
 **/
public class ExcelTools {

    private ExcelTools(){

    }


    private static Logger log = LoggerFactory.getLogger(ExcelTools.class);

    private static final String FILESUFF = ".xlsx";

    private static final String FILESUFFX =".xls";

    /**
     * 读取指定路径下的Excel文件
     * @param filePath
     * @return
     */
    public static Workbook readExcel(String filePath) throws IOException {
        Workbook workbook = null;
        if(filePath == null){
            return null;
        }

        String suffix = filePath.substring(filePath.lastIndexOf("."));
        InputStream in = null;

        try {
            in = ExcelTools.class.getResourceAsStream(filePath);
            if(in == null) {
                in = new FileInputStream(filePath);
            }
            if(suffix.endsWith(FILESUFF)){
                workbook = new XSSFWorkbook(in);
            } else {
                workbook = new HSSFWorkbook(in);
            }

        } catch (FileNotFoundException e){
            log.error("文件名称:{} 不存在",filePath,e);
        } catch (IOException e){
            log.error("解析excel文件:{} 失败",filePath,e);
        } finally {
            if(in != null){
                in.close();
            }
        }
        return workbook;
    }


    /**
     *  读取上传的文件
     * @param file
     * @return
     */
    public static Workbook readExcel(MultipartFile file) throws IOException {
        Workbook wb = null;
        if(!validateExtensionName(file)){
            return null;
        }
        String orginFilePath = file.getOriginalFilename();
        String suffix = orginFilePath.substring(orginFilePath.lastIndexOf("."));

        InputStream is = null;
        try{
            is = file.getInputStream();
            if(suffix.endsWith(FILESUFF)){
                wb = new XSSFWorkbook();
            } else {
                wb = new HSSFWorkbook();
            }

        } catch (FileNotFoundException e){
            log.error("文件名称:{} 不存在",orginFilePath,e);
        } catch (IOException e){
            log.error("解析excel文件:{} 失败",orginFilePath,e);
        } finally {
            if(is != null){
                is.close();
            }
        }
        return wb;
    }

    /**
     * 验证文件 扩展名称是否为 有效的excel文件格式
     * @param file
     * @return
     */
    public static boolean validateExtensionName(MultipartFile file){
        boolean flag = true;
        if(file == null){
            flag =  false;
        }else{
            String originalFilename = file.getOriginalFilename();
            if(StringUtils.isEmpty(originalFilename)){
                flag = false;
            }else {
                String suffix = originalFilename.substring(originalFilename.lastIndexOf("."));
                if(!suffix.equals(FILESUFF) || !suffix.equals(FILESUFFX)){
                    flag = false;
                }
            }
        }
        return flag;
    }
    /**
     * 根据索引获取EXCEL定位符
     * @param columnIndex
     * @param rowIndex
     * @return
     */
    public static String excelColIndexToStr(int columnIndex,int rowIndex) {
        if (columnIndex < 0) {
            return null;
        }
        String columnStr = "";
        rowIndex++;
        do {
            if (columnStr.length() > 0) {
                columnIndex--;
            }
            columnStr = ((char) (columnIndex % 26 + (int) 'A')) + columnStr;
            columnIndex = ((columnIndex - columnIndex % 26) / 26);
        } while (columnIndex > 0);
        return columnStr + rowIndex;
    }
}
