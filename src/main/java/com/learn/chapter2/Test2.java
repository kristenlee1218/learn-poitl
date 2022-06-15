package com.learn.chapter2;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.data.PictureRenderData;
import com.deepoove.poi.util.BytePictureUtils;

import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.HashMap;

/**
 * @author ：Kristen
 * @date ：2022/6/7
 * @description : pot-tl 操作图片
 */
public class Test2 {
    public static void main(String[] args) throws IOException {
        // 编译模板、渲染数据
        XWPFTemplate template = XWPFTemplate.compile("src/main/resources/charpter2/template2.docx").render(new HashMap<String, Object>() {{
            // 本地图片
            put("local", new PictureRenderData(80, 100, "src/main/resources/charpter2/1.png"));
            // 图片流
            put("localbyte", new PictureRenderData(80, 100, ".png", Files.newInputStream(Paths.get("src/main/resources/charpter2/2.png"))));
            // 网络图片(注意网络耗时对系统可能的性能影响)
            put("urlpicture", new PictureRenderData(50, 50, ".png", BytePictureUtils.getUrlBufferedImage("http://deepoove.com/images/icecream.png")));
            // java 图片
            // put("bufferimage", new PictureRenderData(80, 100, ".png", bufferImage)));
        }});
        FileOutputStream out = new FileOutputStream("src/main/resources/charpter2/template2_out.docx");
        // 输出到流
        template.write(out);
        out.flush();
        out.close();
        template.close();
    }
}
