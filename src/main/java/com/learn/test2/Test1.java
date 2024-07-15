package com.learn.test2;

import com.deepoove.poi.XWPFTemplate;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;

public class Test1 {
    public static void main(String[] args) throws IOException {
        // 编译模板、渲染数据
        XWPFTemplate template = XWPFTemplate.compile("D:\\test-poitl\\test.docx").render(new HashMap<String, Object>() {{
            put("rate_leader21_3C_7", "99.5%");
        }});
        FileOutputStream out = new FileOutputStream("D:\\test-poitl\\test_out.docx");
        // 输出到流
        template.write(out);
        out.flush();
        out.close();
        template.close();
    }
}