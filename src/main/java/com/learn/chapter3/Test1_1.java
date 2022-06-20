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
public class Test1_1 {
    public static void main(String[] args) throws IOException {
        HashMap<String, Object> map = new HashMap<>();
        map.put("img", new PictureRenderData(0, 0, ".png", Files.newInputStream(Paths.get("D:\\test-poitl\\2.png"))));
        // 编译模板、渲染数据
        XWPFTemplate template = XWPFTemplate.compile("D:\\test-poitl\\template.docx").render(map);
        FileOutputStream out = new FileOutputStream("D:\\test-poitl\\output.docx");
        // 输出到流
        template.write(out);
        out.flush();
        out.close();
        template.close();
    }
}
