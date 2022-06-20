package com.learn.chapter2;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.data.NumbericRenderData;
import com.deepoove.poi.data.TextRenderData;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;

/**
 * @author ：Kristen
 * @date ：2022/6/15
 * @description :
 */
public class Test4_1 {
    public static void main(String[] args) throws IOException {
        XWPFTemplate template = XWPFTemplate.compile("src/main/resources/charpter2/template4.docx")
                .render(new HashMap<String, Object>() {{
                    put("list", new NumbericRenderData(new ArrayList<TextRenderData>() {
                        {
                            add(new TextRenderData("Plug-in function, define your own function"));
                            add(new TextRenderData("Supports word text, header..."));
                            add(new TextRenderData("Not just templates"));
                        }
                    }));
                }});
        FileOutputStream out;
        out = new FileOutputStream("src/main/resources/charpter2/template4_out.docx");
        template.write(out);
        out.flush();
        out.close();
        template.close();
    }
}
