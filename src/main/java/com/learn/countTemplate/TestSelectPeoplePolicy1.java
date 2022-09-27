package com.learn.countTemplate;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import com.deepoove.poi.config.ConfigureBuilder;
import com.learn.renyuanTemplate.ListQuotaTablePolicy;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;

/**
 * @author ：Kristen
 * @date ：2022/9/27
 * @description :
 */
public class TestSelectPeoplePolicy1 {
    public static void main(String[] args) throws IOException {
        // 准备数据
        ConfigureBuilder builder = Configure.newBuilder();
        // 编译模板
        Configure configure = builder.build();
        HashMap<String, Object> data = new HashMap<String, Object>() {
            {
                // 构建 table
                SelectPeoplePolicy1 policy = new SelectPeoplePolicy1();
                builder.bind("table", policy);
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
