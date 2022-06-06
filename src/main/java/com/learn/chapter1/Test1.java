package com.learn.chapter1;

import com.deepoove.poi.XWPFTemplate;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;

/**
 * @author ：Kristen
 * @date ：2022/6/6
 * @description :
 */
public class Test1 {
    public static void main(String[] args) throws IOException {
        XWPFTemplate template = XWPFTemplate.compile("D:\\template.docx").render(
                new HashMap<String, Object>() {{
                    put("title", "Hi, poi-tl Word模板引擎");
                }});
        FileOutputStream out = new FileOutputStream("D:\\output.docx");
        template.write(out);
        out.flush();
        out.close();
        template.close();
    }
}
