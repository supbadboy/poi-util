package com.adingxiong.poiutils.code;

import com.adingxiong.poiutils.constant.Errorcons;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageSz;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STPageOrientation;
import org.springframework.util.Assert;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.math.BigInteger;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;
import java.util.stream.Stream;

/**
 * @ClassName WordExport
 * @Description TODO
 * @Author xiongchao
 * @Date 2020/12/1 10:20
 **/
public class WordExport {

    public static WordExport getInstance() {
        return new WordExport();
    }

    public Path convertWord(InputStream in, Map<String, Object> params, String fileName) throws IOException {
        Assert.notNull(in, Errorcons.PARM_EMPTY);
        Assert.notNull(fileName, Errorcons.PARM_EMPTY);
        Assert.notNull(params, Errorcons.PARM_EMPTY);
        Map<String, Object> newParams = new HashMap();
        params.forEach((k, v) -> {
            newParams.put("${" + k + "}", v);
        });
        Path path = Paths.get(System.getProperty("java.io.tmpdir"), fileName);

        try {
            OutputStream out = Files.newOutputStream(path);
            Throwable var7 = null;

            try {
                XWPFDocument doc = new XWPFDocument(in);
                this.getParamTable(doc, params, false);
                this.replaceInPara((XWPFDocument)doc, newParams);
                this.replaceInTable(doc, newParams);
                doc.write(out);
            } catch (Throwable var17) {
                var7 = var17;
                throw var17;
            } finally {
                this.close(out);
            }
        } catch (Exception var19) {
            var19.printStackTrace();
        }

        return path;
    }

    public Path convertWord(InputStream in, Map<String, Object> params, String fileName, boolean horizontal) throws IOException {
        Assert.notNull(in, Errorcons.PARM_EMPTY);
        Assert.notNull(fileName, Errorcons.PARM_EMPTY);
        Assert.notNull(params, Errorcons.PARM_EMPTY);
        Map<String, Object> newParams = new HashMap();
        params.forEach((k, v) -> {
            newParams.put("${" + k + "}", v);
        });
        Path path = Paths.get(System.getProperty("java.io.tmpdir"), fileName);

        try {
            OutputStream out = Files.newOutputStream(path);
            Throwable var8 = null;

            try {
                XWPFDocument doc = new XWPFDocument(in);
                this.getParamTable(doc, params, horizontal);
                this.replaceInPara((XWPFDocument)doc, newParams);
                this.replaceInTable(doc, newParams);
                doc.write(out);
            } catch (Throwable var18) {
                var8 = var18;
                throw var18;
            } finally {
                this.close(out);

            }
        } catch (Exception var20) {
            var20.printStackTrace();
        }

        return path;
    }

    private WordExport() {
    }

    /**
     *  获取list文件 并渲染到页面
     * @param doc
     * @param params
     * @param horizontal  是否自动调整格式
     * @throws Exception
     */
    public void getParamTable(XWPFDocument doc, Map<String, Object> params, boolean horizontal) throws Exception {
        List<Object> list = null;
        Set<String> checkKey = new HashSet();
        Map<String, Object> resVal = new HashMap();
        Iterator var7 = params.entrySet().iterator();

        while(true) {
            Map.Entry entry;
            String objType;
            do {
                do {
                    if (!var7.hasNext()) {
                        if (list == null) {
                            return;
                        }

                        if (list.isEmpty()) {
                            return;
                        }

                        XWPFTable table = (XWPFTable)doc.getTables().get(0);
                        if (horizontal) {
                            CTSectPr ctSectPr = doc.getDocument().getBody().getSectPr();
                            CTPageSz pgSz = ctSectPr.getPgSz();
                            pgSz.setW(BigInteger.valueOf(16840L));
                            pgSz.setH(BigInteger.valueOf(11907L));
                            pgSz.setOrient(STPageOrientation.LANDSCAPE);
                        }

                        List<XWPFTableRow> rows = table.getRows();
                        int i = 0;
                        XWPFTableRow ins = null;
                        Map<String, Integer> cellMap = new HashMap();
                        Iterator var27 = rows.iterator();

                        while(var27.hasNext()) {
                            XWPFTableRow row = (XWPFTableRow)var27.next();
                            List<XWPFTableCell> tableCells = row.getTableCells();
                            int j = 0;

                            for(Iterator var16 = tableCells.iterator(); var16.hasNext(); ++j) {
                                XWPFTableCell cell = (XWPFTableCell)var16.next();
                                String text = (String)cell.getParagraphs().stream().map(XWPFParagraph::getRuns).flatMap(Stream::of).map((x) -> {
                                    return x.toString();
                                }).collect(Collectors.joining());
                                String reText = text.replaceAll(" ", "").replaceAll(",", "").replace("[", "").replace("]", "").replaceAll("\n", "").trim();
                                if (checkKey.contains(reText)) {
                                    if (ins == null) {
                                        ins = row;
                                    }

                                    cellMap.put(reText, j);
                                }
                            }

                            if (ins == null) {
                                ++i;
                            }
                        }

                        for(int l = 0; l < list.size(); ++l) {
                            Iterator var30 = cellMap.entrySet().iterator();

                            while(var30.hasNext()) {
                                Map.Entry<String, Integer> entrys = (Map.Entry)var30.next();
                                ins.getCell((Integer)entrys.getValue()).removeParagraph(0);
                                ins.getCell((Integer)entrys.getValue()).addParagraph().createRun().setText(String.valueOf(resVal.get(l + "." + (String)entrys.getKey())));
                            }

                            table.addRow(ins, i + l);
                        }

                        table.removeRow(i + list.size() - 1);
                        return;
                    }

                    entry = (Map.Entry)var7.next();
                } while(entry.getValue() == null);

                objType = entry.getValue().getClass().getSimpleName();
            } while(!"ArrayList".equals(objType));

            list = (List)entry.getValue();
            int i = 0;

            for(Iterator var11 = list.iterator(); var11.hasNext(); ++i) {
                Object obj = var11.next();
                Map<String, Object> map = getObjectToMap(obj);
                Iterator var14 = map.entrySet().iterator();

                while(var14.hasNext()) {
                    Map.Entry<String, Object> item = (Map.Entry)var14.next();
                    checkKey.add("${" + (String)entry.getKey() + "." + (String)item.getKey() + "}");
                    resVal.put(i + ".${" + (String)entry.getKey() + "." + (String)item.getKey() + "}", item.getValue());
                }
            }
        }
    }

    public static Map<String, Object> getObjectToMap(Object obj) throws IllegalAccessException {
        Map<String, Object> map = new LinkedHashMap();
        Class<?> clazz = obj.getClass();
        Field[] var3 = clazz.getDeclaredFields();
        int var4 = var3.length;

        for(int var5 = 0; var5 < var4; ++var5) {
            Field field = var3[var5];
            field.setAccessible(true);
            String fieldName = field.getName();
            Object value = field.get(obj);
            if (value == null) {
                value = "";
            }

            map.put(fieldName, value);
        }

        return map;
    }

    private void replaceInPara(XWPFDocument doc, Map<String, Object> params) {
        Iterator iterator = doc.getParagraphsIterator();

        while(iterator.hasNext()) {
            XWPFParagraph para = (XWPFParagraph)iterator.next();
            this.replaceInPara(para, params);
        }

    }

    private void replaceInPara(XWPFParagraph para, Map<String, Object> params) {
        if (this.matcher(para.getParagraphText()).find()) {
            List<XWPFRun> runs = para.getRuns();
            int start = -1;
            int end = -1;
            String str = "";
            String color = null;
            int fontSize = 0;
            UnderlinePatterns underlinePatterns = null;

            int i;
            String runText;
            for(i = 0; i < runs.size(); ++i) {
                XWPFRun run = (XWPFRun)runs.get(i);
                if (fontSize == 0) {
                    fontSize = run.getFontSize();
                }

                if (underlinePatterns == null) {
                    underlinePatterns = run.getUnderline();
                }

                if (color == null) {
                    color = run.getColor();
                }

                runText = run.toString();
                if (runText.length() == 1 && '$' == runText.charAt(0)) {
                    start = i;
                } else if (runText.length() > 1 && '$' == runText.charAt(0) && '{' == runText.charAt(1)) {
                    start = i;
                }

                if (start != -1) {
                    str = str + runText;
                }

                if ('}' == runText.charAt(runText.length() - 1) && start != -1) {
                    end = i;
                    break;
                }
            }

            for(i = start; i <= end; ++i) {
                para.removeRun(i);
                --i;
                --end;
            }

            Iterator var18 = params.keySet().iterator();

            while(var18.hasNext()) {
                String key = (String)var18.next();
                if (str.equals(key)) {
                    runText = null;
                    if (params.get(key) != null) {
                        runText = params.get(key).getClass().getSimpleName();
                    }

                    XWPFRun run = para.createRun();
                    if (!"InputStream".equals(runText) && !"HttpInputStream".equals(runText)) {
                        if (params.get(key) != null) {
                            run.setFontSize(fontSize);
                            run.setColor(color);
                            run.setUnderline(underlinePatterns);
                            run.setText(String.valueOf(params.get(key)), 0);
                        }
                    } else {
                        try {
                            InputStream inputStream = (InputStream)params.get(key);
                            run.addPicture(inputStream, 5, "qianming.jpg", Units.toEMU(100.0D), Units.toEMU(60.0D));
                            inputStream.close();
                        } catch (Exception var16) {
                            var16.printStackTrace();
                        }
                    }
                    break;
                }
            }
        }

    }

    private void replaceInPara1(XWPFParagraph para, Map<String, Object> params) {
        String runText = "";
        String color = null;
        int fontSize = 0;
        UnderlinePatterns underlinePatterns = null;
        if (this.matcher(para.getParagraphText()).find()) {
            List<XWPFRun> runs = para.getRuns();
            if (runs.size() > 0) {
                int j = runs.size();

                for(int i = 0; i < j; ++i) {
                    XWPFRun run = (XWPFRun)runs.get(0);
                    String i1 = run.toString();
                    runText = runText + i1;
                    para.removeRun(0);
                }
            }

            Matcher matcher = this.matcher(runText);
            if (matcher.find()) {
                while((matcher = this.matcher(runText)).find()) {
                    runText = matcher.replaceFirst(String.valueOf(params.get(matcher.group(1))));
                }

                XWPFRun run = para.createRun();
                run.setText(runText, 0);
                run.setFontSize(fontSize);
                run.setUnderline((UnderlinePatterns)underlinePatterns);
            }
        }

    }

    private void replaceInTable(XWPFDocument doc, Map<String, Object> params) {
        Iterator iterator = doc.getTablesIterator();

        while(iterator.hasNext()) {
            XWPFTable table = (XWPFTable)iterator.next();
            List<XWPFTableRow> rows = table.getRows();
            Iterator var8 = rows.iterator();

            while(var8.hasNext()) {
                XWPFTableRow row = (XWPFTableRow)var8.next();
                List<XWPFTableCell> cells = row.getTableCells();
                Iterator var10 = cells.iterator();

                while(var10.hasNext()) {
                    XWPFTableCell cell = (XWPFTableCell)var10.next();
                    List<XWPFParagraph> paras = cell.getParagraphs();
                    Iterator var12 = paras.iterator();

                    while(var12.hasNext()) {
                        XWPFParagraph para = (XWPFParagraph)var12.next();
                        this.replaceInPara(para, params);
                    }
                }
            }
        }

    }

    private Matcher matcher(String str) {
        Pattern pattern = Pattern.compile("\\$\\{(.+?)\\}", 2);
        Matcher matcher = pattern.matcher(str);
        return matcher;
    }

    private void close(InputStream is) {
        if (is != null) {
            try {
                is.close();
            } catch (IOException var3) {
                var3.printStackTrace();
            }
        }

    }

    private void close(OutputStream os) {
        if (os != null) {
            try {
                os.close();
            } catch (IOException var3) {
                var3.printStackTrace();
            }
        }

    }


}
