package com.gitee.poiutils.handler;




import com.gitee.poiutils.util.ExcelCell;

import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.ConcurrentMap;

/**
 * @ClassName DefHandler
 * @Description TODO
 * @Author xiongchao
 * @Date 2020/10/21 13:22
 **/
public interface DefHandler {

    ConcurrentMap<String,DefHandler> handlers = new ConcurrentHashMap<String, DefHandler>();

    String excute(ExcelCell excelCell, Map<String, String> param);
}
