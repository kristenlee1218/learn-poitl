package com.learn.test;

import com.deepoove.poi.XWPFTemplate;

import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

/**
 * @author ：Kristen
 * @date ：2022/6/20
 * @description :
 */
public class Test1 {
    public static void main(String[] args) throws IOException {
        Map<String, Object> map = new HashMap<>();

        // sheet所需要的信息
        Data data = new Data();
        data.setDate("2022");
        data.setCompany("中国兵器工业信息中心");
        map.put("data", data);

        // 分值表格需要的信息
        Type type1 = new Type();
        type1.setA1A2A3("72.54");
        type1.setB("79.09");
        type1.setC("69.77 ");
        type1.setSubCount("76.23");
        type1.setD("68.80");
        type1.setCount("75.12");
        map.put("type1", type1);

        Group group1 = new Group();
        group1.setA1A2A3("70.45");
        group1.setB("76.79");
        group1.setC("70.93");
        group1.setSubCount("75.48");
        group1.setD("69.84");
        group1.setCount("74.63");
        map.put("group1", group1);

        XWPFTemplate template = XWPFTemplate.compile("src/main/resources/test/template2.docx").render(map);
        template.writeToFile("src/main/resources/test/template2_out.docx");
    }
}
