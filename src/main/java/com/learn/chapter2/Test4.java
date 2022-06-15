package com.learn.chapter2;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.data.NumbericRenderData;
import com.deepoove.poi.data.TextRenderData;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;

/**
 * @author ：Kristen
 * @date ：2022/6/7
 * @description : poi-tl 操作列表
 */
public class Test4 {
    public static void main(String[] args) throws IOException {
        Map<String, Object> data = new HashMap<>();
        data.put("list", new NumbericRenderData(new ArrayList<TextRenderData>() {
            {
                add(new TextRenderData("Plug-in function, define your own function"));
                add(new TextRenderData("Supports word text, header..."));
                add(new TextRenderData("Not just templates"));
            }
        }));
        XWPFTemplate template = XWPFTemplate.compile("src/main/resources/charpter2/template4.docx")
                .render(data);
        FileOutputStream out;
        out = new FileOutputStream("src/main/resources/charpter2/template4.docx");
        template.write(out);
        out.flush();
        out.close();
        template.close();
    }
}
