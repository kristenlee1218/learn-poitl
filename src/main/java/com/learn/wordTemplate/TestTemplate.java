package com.learn.wordTemplate;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import com.deepoove.poi.config.ConfigureBuilder;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;

public class TestTemplate {
    public static void main(String[] args) throws IOException {
        //准备数据
        ConfigureBuilder builder = Configure.newBuilder();
        //编译模板
        Configure configure = builder.build();
        HashMap<String, Object> data = new HashMap<String, Object>() {
            {
                // 构建table
                TestPolicy testPolicy = new TestPolicy();
                builder.bind("table", testPolicy);
            }
        };

        XWPFTemplate template = XWPFTemplate.compile("D:\\test-poitl\\ele02.docx", configure).render(data);
        FileOutputStream out = new FileOutputStream("D:\\test-poitl\\ele02_out.docx");
        template.write(out);
        out.flush();
        out.close();
        template.close();
    }
}