package com.learn.chapter2;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.data.MiniTableRenderData;
import com.deepoove.poi.data.NumbericRenderData;
import com.deepoove.poi.data.RowRenderData;
import com.deepoove.poi.data.TextRenderData;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Map;

/**
 * @author ：Kristen
 * @date ：2022/6/7
 * @description :
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
        XWPFTemplate template = XWPFTemplate.compile("D:\\test-poitl\\template.docx")
                .render(data);
        FileOutputStream out;
        out = new FileOutputStream("D:\\test-poitl\\output.docx");
        template.write(out);
        out.flush();
        out.close();
        template.close();
    }
}
