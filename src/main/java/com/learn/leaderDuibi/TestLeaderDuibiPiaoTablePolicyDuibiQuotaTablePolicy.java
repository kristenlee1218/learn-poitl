package com.learn.leaderDuibi;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import com.deepoove.poi.config.ConfigureBuilder;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;

/**
 * @author ：Kristen
 * @date ：2023/8/15
 * @description :
 */
public class TestLeaderDuibiPiaoTablePolicyDuibiQuotaTablePolicy {
    public static void main(String[] args) throws IOException {
        // 准备数据
        ConfigureBuilder builder = Configure.newBuilder();
        // 编译模板
        Configure configure = builder.build();
        HashMap<String, Object> data = new HashMap<String, Object>() {
            {
                // 构建 table
                LeaderDuibiPiaoTablePolicyDuibiQuotaTablePolicy tablePolicy = new LeaderDuibiPiaoTablePolicyDuibiQuotaTablePolicy();
                builder.bind("table", tablePolicy);
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
