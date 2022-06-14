package com.learn.example5;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.data.PictureRenderData;
import org.junit.jupiter.api.Test;

import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.HashMap;
import java.util.Map;

/**
 * @author ：Kristen
 * @date ：2022/6/14
 * @description : 证书
 */
public class CertificateExample {

    @SuppressWarnings("serial")
    @Test
    public void testRenderTextBox() throws Exception {
        Map<String, Object> data = new HashMap<String, Object>() {
            {
                put("name", "Poi-tl");
                put("department", "DEEPOOVE.COM");
                put("y", "2020");
                put("m", "8");
                put("d", "19");
                put("img", new PictureRenderData(120, 120, ".png", Files.newInputStream(Paths.get("D:\\test-poitl\\lu.png"))));
            }
        };

        XWPFTemplate.compile("D:\\test-poitl\\template5.docx").render(data).writeToFile("D:\\test-poitl\\template5_out.docx");

    }
}