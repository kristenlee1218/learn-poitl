package com.learn.chapter3;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.data.PictureRenderData;

import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.HashMap;

/**
 * @author ：Kristen
 * @date ：2022/6/7
 * @description : poi-tl 操作图片
 */
public class Test1 {
    public static void main(String[] args) throws IOException {
        // 编译模板、渲染数据
        XWPFTemplate template = XWPFTemplate.compile("src/main/resources/charpter3/template1.docx").render(
                new HashMap<String, Object>() {{
                    put("img", new PictureRenderData(0, 0, ".png", Files.newInputStream(Paths.get("D:\\test-poitl\\2.png"))));
                }});
        FileOutputStream out = new FileOutputStream("src/main/resources/charpter3/template1_out.docx");
        // 输出到流
        template.write(out);
        out.flush();
        out.close();
        template.close();
    }
}
