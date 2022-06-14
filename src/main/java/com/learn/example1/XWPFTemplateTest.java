package com.learn.example1;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.data.*;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

/**
 * @author ：Kristen
 * @date ：2022/6/13
 * @description : 软件说明文档
 */
public class XWPFTemplateTest {
    public static void main(String[] args) throws IOException {
        final RowRenderData header = new RowRenderData(
                Arrays.asList(new TextRenderData("FFFFFF", "Word处理解决方案"),
                        new TextRenderData("FFFFFF", "是否跨平台"), new TextRenderData("FFFFFF", "易用性")),
                "ff9800");
        RowRenderData row0 = RowRenderData.build("Poi-tl", "纯Java组件，跨平台", "简单：模板引擎功能，并对POI进行了一些封装");
        RowRenderData row1 = RowRenderData.build("Apache Poi", "纯Java组件，跨平台", "简单，缺少一些功能的封装");
        RowRenderData row2 = RowRenderData.build("Freemarker", "XML操作，跨平台", "复杂，需要理解XML结构");
        RowRenderData row3 = RowRenderData.build("OpenOffice", "需要安装OpenOffice软件",
                "复杂，需要了解OpenOffice的API");
        RowRenderData row4 = RowRenderData.build("Jacob、winlib", "Windows平台", "复杂，不推荐使用");
        final List<RowRenderData> tableData = Arrays.asList(row0, row1, row2, row3, row4);
        Map<String, Object> data = new HashMap<String, Object>() {
            {
                put("header", "Deeply love what you love.");
                put("name", "Poi-tl");
                put("word", "模板引擎");
                put("time", "2019-05-31");
                put("what",
                        "Java Word模板引擎： Minimal Microsoft word(docx) templating with {{template}} in Java. It works by expanding tags in a template using values provided in a JavaMap or JavaObject.");
                put("author", new TextRenderData("000000", "Sayi卅一"));
                put("introduce", new HyperLinkTextRenderData("http://www.deepoove.com", "http://www.deepoove.com"));
                put("portrait", new PictureRenderData(60, 60, "src/test/resources/sayi.png"));

                put("solution_compare", new MiniTableRenderData(header, tableData,
                        MiniTableRenderData.WIDTH_A4_FULL));

                put("feature", new NumbericRenderData(new ArrayList<TextRenderData>() {
                    {
                        add(new TextRenderData("Plug-in grammar, add new grammar by yourself"));
                        add(new TextRenderData("Supports word text, local pictures, web pictures, table, list, header, footer..."));
                        add(new TextRenderData("Templates, not just templates, but also style templates"));
                    }
                }));
            }
        };

        XWPFTemplate template = XWPFTemplate.compile("D:\\test-poitl\\template1.docx")
                .render(data);
        FileOutputStream out = new FileOutputStream("D:\\test-poitl\\out_template1.docx");
        template.write(out);
        out.flush();
        out.close();
        template.close();
    }
}
