package com.learn.statictisDetail;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import com.deepoove.poi.config.ConfigureBuilder;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;

/**
 * @author ：Kristen
 * @date ：2023/8/21
 * @description :
 */
public class TestStatisticDetailTablePolicy {
    public static void main(String[] args) throws IOException {
        // 准备数据
        ConfigureBuilder builder = Configure.newBuilder();
        // 编译模板
        Configure configure = builder.build();
        HashMap<String, Object> data = new HashMap<String, Object>() {
            {
                // 构建 table
                StatisticDetailTablePolicy statisticDetailTablePolicy = new StatisticDetailTablePolicy();
                builder.bind("table", statisticDetailTablePolicy);
            }
        };

        XWPFTemplate template = XWPFTemplate.compile("D:\\test-poitl\\detail.docx", configure).render(data);
        FileOutputStream out = new FileOutputStream("D:\\test-poitl\\detail_out.docx");
        template.write(out);
        out.flush();
        out.close();
    }
}