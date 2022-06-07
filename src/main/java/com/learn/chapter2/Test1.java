package com.learn.chapter2;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.data.HyperLinkTextRenderData;
import com.deepoove.poi.data.TextRenderData;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;

/**
 * @author ：Kristen
 * @date ：2022/6/7
 * @description : poi-tl 操作文字等
 */
public class Test1 {
    public static void main(String[] args) throws IOException {
        // 编译模板、渲染数据
        XWPFTemplate template = XWPFTemplate.compile("D:\\test-poitl\\template.docx").render(
                new HashMap<String, Object>() {{
                    put("name", "Sayi");
                    put("author", new TextRenderData("000000", "Sayi"));
                    put("link", new HyperLinkTextRenderData("website", "http://deepoove.com"));
                    put("anchor", new HyperLinkTextRenderData("anchortxt", "anchor:appendix1"));
                }});
        FileOutputStream out = new FileOutputStream("D:\\test-poitl\\output.docx");
        // 输出到流
        template.write(out);
        out.flush();
        out.close();
        template.close();
    }
}
