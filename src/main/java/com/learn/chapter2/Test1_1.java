package com.learn.chapter2;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.data.HyperLinkTextRenderData;
import com.deepoove.poi.data.TextRenderData;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

/**
 * @author ：Kristen
 * @date ：2022/6/15
 * @description :
 */
public class Test1_1 {
    public static void main(String[] args) throws IOException {
        Map<String, Object> map = new HashMap<>();
        map.put("name", "Sayi");
        map.put("author", new TextRenderData("000000", "Sayi"));
        map.put("link", new HyperLinkTextRenderData("website", "http://deepoove.com"));
        map.put("anchor", new HyperLinkTextRenderData("anchortxt", "anchor:appendix1"));
        // 编译模板、渲染数据
        XWPFTemplate template = XWPFTemplate.compile("src/main/resources/charpter2/template1.docx").render(map);
        FileOutputStream out = new FileOutputStream("src/main/resources/charpter2/template1_out.docx");
        // 输出到流
        template.write(out);
        out.flush();
        out.close();
        template.close();
    }
}
