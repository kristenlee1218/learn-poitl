package com.learn.chapter2;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.data.PictureRenderData;
import com.deepoove.poi.util.BytePictureUtils;

import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.HashMap;
import java.util.Map;

/**
 * @author ：Kristen
 * @date ：2022/6/15
 * @description :
 */
public class Test2_1 {
    public static void main(String[] args) throws IOException {
        Map<String, Object> map = new HashMap<>();
        map.put("local", new PictureRenderData(80, 100, "src/main/resources/charpter2/1.png"));
        map.put("localbyte", new PictureRenderData(80, 100, ".png", Files.newInputStream(Paths.get("src/main/resources/charpter2/2.png"))));
        map.put("urlpicture", new PictureRenderData(50, 50, ".png", BytePictureUtils.getUrlBufferedImage("http://deepoove.com/images/icecream.png")));
        // 编译模板、渲染数据
        XWPFTemplate template = XWPFTemplate.compile("src/main/resources/charpter2/template2.docx").render(map);
        FileOutputStream out = new FileOutputStream("src/main/resources/charpter2/template2_out.docx");
        // 输出到流
        template.write(out);
        out.flush();
        out.close();
        template.close();
    }
}
