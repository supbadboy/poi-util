##   `poi-util` 加强`POI`相关操作工具包

### 迭代记录  
+ 2020年11月27日11:33:35 完成Excel导出的部分


### 工具包说明 
| 继承工具类        |            | 完结时间  |
| ------------- |:-------------:| -----:|
| Excel导出      |  |  |
|    |  通用导出`fillCustomer`  |  2020年11月27日11:33:35  |
|   |定制导出  `fillCommon`  |    2020年11月27日11:33:35 |
|Excel 导入  |    |  开发中    |
|   Excel转HTML    |       | 开发中      |
|  WORD模版导出     |       |    开发中   |
|  Excel 模版导出     |       |  未开发     |
|       |       |       |

#### 1. 结构说明
![](docs/Snipaste_2020-11-27_11-36-54.png)

#### 2. 使用说明 

    该部分的核心功能是自定义注解，在你需要导出的实体类中，
    将你需要作为表头导出的字段加上自定义注解 @FieldName ，
    注解的属性有value，为Excel表头名称 require是否必填，默认为否
    simpleDate时间格式为，默认为 空  



#### 3. 调用说明
实体类示例 
```java
@Data
public class ProjectVo {

    @FieldName(value = "项目")
    private String name;

    @FieldName(value = "电话")
    private String phone;

    private String person;

    @FieldName(value = "金额")
    private Double money;

    @FieldName(value = "负责人")
    private String processPeople;

    @FieldName(value = "周期")
    private String cycle;

    @FieldName(value = "记录日期",dateFormat = "yyyy-MM-dd")
    private Date date;
}

```

导出功能调用,目前`Excel`导出分两种

 + 通用导出(默认注解value为表头,根据字段值的长度来自适应宽度,全局字体上下左右居中)
 
 + 定制导出 (支持多sheet页导出,支持自定义标题,表头,样式,字体等)
 
 
 ##### 3.1  通用导出`Excel`
 调用方法 
```
    
    Workbook wb = new XSSFWorkbook();
    //模拟数据
    List<ProjectVo> list = mockData();
    //调用导出方法 
    // 参数说明 wb对象   集合数据  sheetName 默认为sheet
    ExcelExport.getInstance().fillCommon(wb,list,null);
    FileOutputStream out = new FileOutputStream("C:\\Users\\Administrator\\Desktop\\测试导出.xlsx");
    wb.write(out);
    out.close();
```

效果如下 

![](docs/Snipaste_2020-11-27_13-32-12.png)

##### 3.2 定制导出`Excel`
```
    //定制excel
    Workbook wb = new XSSFWorkbook();
    Sheet sheet = wb.createSheet("定制表头导出");
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
    
    
    //核心导出部分  前面部分可以自己定制表头,标题,全局样式,多sheet等 
    
    ExcelExport.getInstance().fillCustomer(sheet,1,null,cellStyle,list);
    FileOutputStream out = new FileOutputStream("C:\\Users\\Administrator\\Desktop\\定制导出.xlsx");
    wb.write(out);
    out.close();
```

效果如下所示 
![](docs/Snipaste_2020-11-27_13-30-18.png)

![](docs/Snipaste_2020-11-27_13-30-30.png)
