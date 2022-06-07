package com.learn.chapter2;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.data.MiniTableRenderData;
import com.deepoove.poi.data.RowRenderData;
import com.deepoove.poi.data.TextRenderData;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Map;

/**
 * @author ：Kristen
 * @date ：2022/6/7
 * @description : pot-tl 操作表格
 */
public class Test3 {
    public static void main(String[] args) throws IOException {
        Map<String, Object> data = new HashMap<>();
        RowRenderData header = RowRenderData.build(new TextRenderData("FF0000", "姓名"), new TextRenderData("FF0000", "学历"));

        RowRenderData row0 = RowRenderData.build("张三", "研究生");
        RowRenderData row1 = RowRenderData.build("李四", "博士");

        data.put("table", new MiniTableRenderData(header, Arrays.asList(row0, row1)));
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
