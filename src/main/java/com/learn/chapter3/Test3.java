package com.learn.chapter3;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.data.ChartSingleSeriesRenderData;
import com.deepoove.poi.data.SeriesRenderData;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;

/**
 * @author ：Kristen
 * @date ：2022/6/7
 * @description : poi-tl 操作单图表
 */
public class Test3 {
    public static void main(String[] args) throws IOException {
        // 编译模板、渲染数据
        XWPFTemplate template = XWPFTemplate.compile("D:\\test-poitl\\template.docx").render(
                new HashMap<String, Object>() {{
                    ChartSingleSeriesRenderData pie = new ChartSingleSeriesRenderData();
                    pie.setChartTitle("ChartTitle");
                    pie.setCategories(new String[]{"俄罗斯", "加拿大", "美国", "中国"});
                    pie.setSeriesData(new SeriesRenderData("countries", new Integer[]{17098242, 9984670, 9826675, 9596961}));
                    put("pieChart", pie);
                }});
        FileOutputStream out = new FileOutputStream("D:\\test-poitl\\output.docx");
        // 输出到流
        template.write(out);
        out.flush();
        out.close();
        template.close();
    }
}
