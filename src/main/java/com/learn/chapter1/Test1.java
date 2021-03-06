package com.learn.chapter1;

import com.deepoove.poi.XWPFTemplate;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;

/**
 * @author ：Kristen
 * @date ：2022/6/6
 * @description : Template 模板、Data-model 数据、Output 输出
 */
public class Test1 {
    public static void main(String[] args) throws IOException {
        // 编译模板、渲染数据
        XWPFTemplate template = XWPFTemplate.compile("src/main/resources/charpter1/template1.docx").render(new HashMap<String, Object>() {{
            put("title", "Hi, poi-tl Word 模板引擎");
        }});
        FileOutputStream out = new FileOutputStream("src/main/resources/charpter1/template1_out.docx");
        // 输出到流
        template.write(out);
        out.flush();
        out.close();
        template.close();
    }
}
