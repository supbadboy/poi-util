package com.gitee.poiutils.handler;


import com.gitee.poiutils.util.ExcelCell;
import com.greenpineyu.fel.Fel;
import org.apache.commons.lang3.StringUtils;
import java.util.Arrays;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

/**
 * @author Administrator
 * @date 2020/9/9
 * @Description: TODO
 */
public class FormulaHandler implements DefHandler{

    @Override
    public String excute(ExcelCell excelCell, Map<String,String> param) {
        if(excelCell.getFormula() == null){
            return "&nbsp;";
        }
        String formula = excelCell.getFormula();
        String elements = formula
                .replaceAll("\\(", ",")
                .replaceAll("\\)", ",")
                .replaceAll("\\*", ",")
                .replaceAll("\\/", ",")
                .replaceAll("\\-", ",")
                .replaceAll("\\+", ",");
        List<String> strings = Arrays.asList(elements.split(",")).stream()
                .filter(StringUtils::isNotBlank)
                .map(String::trim)
                .collect(Collectors.toList());
        for(String i : strings){
            if(param.containsKey(i)){
                formula = formula.replaceAll(i,param.getOrDefault(i,""));
            }
        }
        try {
            return Fel.newEngine().eval(formula).toString();
        } catch (Exception e) {
            e.printStackTrace();
            return "&nbsp;";
        }
    }

}
