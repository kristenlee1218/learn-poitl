package com.learn.test;

import com.deepoove.poi.XWPFTemplate;

import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author ：Kristen
 * @date ：2022/6/20
 * @description :
 */
public class Test2 {
    public static void main(String[] args) throws IOException {
        Map<String, Object> map = new HashMap<>();

        // sheet所需要的信息
        Data data = new Data();
        data.setDate("2022");
        data.setCompany("中国兵器工业信息中心");
        map.put("data", data);

        // 分值表格需要的信息
        //List<Type> list = new ArrayList<>();
        Type type = new Type();
        type.setA1A2A3("72.54");
        type.setB("79.09");
        type.setC("69.77 ");
        type.setSubCount("76.23");
        type.setD("68.80");
        type.setCount("75.12");
        //list.add(type);
        map.put("type", type);

        XWPFTemplate template = XWPFTemplate.compile("src/main/resources/test/template.docx").render(map);
        template.writeToFile("src/main/resources/test/template_out.docx");
    }
}
