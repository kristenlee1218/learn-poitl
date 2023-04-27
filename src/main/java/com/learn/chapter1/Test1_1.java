package com.learn.chapter1;

import com.deepoove.poi.XWPFTemplate;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

/**
 * @author ：Kristen
 * @date ：2022/6/15
 * @description :
 *
 * 1、Template 模板
 * （1）、模板是 Docx 格式的 Word 文档。所有的标签都是以 “{{” 开头，以 “}}” 结尾，模板标签可以出现在任何位置，包括页眉，页脚，表格内部，文本框等等
 *
 * （2）、poi-tl 遵循 “所见即所得” 的设计，模板的样式会被完全保留，标签的样式也会应用在替换后的文本上
 *
 * 2、
 *
 *
 */
public class Test1_1 {
    public static void main(String[] args) throws IOException {
        Map<String, Object> map = new HashMap<>();
        map.put("title", "Hi, poi-tl Word 模板引擎");
        map.put("content", "这是内容");
        // 编译模板、渲染数据
        XWPFTemplate template = XWPFTemplate.compile("src/main/resources/charpter1/template1.docx").render(map);
        FileOutputStream out = new FileOutputStream("src/main/resources/charpter1/template1_out.docx");
        // 输出到流
        template.write(out);
        out.flush();
        out.close();
        template.close();
    }
}
