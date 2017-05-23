package org.xmx0632.exportfile.excel.model;

import java.util.HashMap;
import java.util.Map;
 
/**
 * Excel数据导入导出格式化<br>
 * 举例:<br>
 * 数据导出， {lock,{0:正常，1:锁定}}<br>
 */
public class ExcelDataFormatter {
     
    private Map<String,Map<String,String>> formatter=new HashMap<String, Map<String,String>>();
 
    public void set(String key,Map<String,String> map){
        formatter.put(key, map);
    }
     
    public Map<String,String> get(String key){
        return formatter.get(key);
    }
     
}
