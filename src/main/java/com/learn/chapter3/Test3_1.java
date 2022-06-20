package com.learn.chapter3;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.data.ChartSingleSeriesRenderData;
import com.deepoove.poi.data.SeriesRenderData;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

/**
 * @author ：Kristen
 * @date ：2022/6/7
 * @description : poi-tl 操作单图表
 */
public class Test3_1 {
    public static void main(String[] args) throws IOException {
        ChartSingleSeriesRenderData pie = new ChartSingleSeriesRenderData();
        pie.setChartTitle("ChartTitle");
        pie.setCategories(new String[]{"俄罗斯", "加拿大", "美国", "中国"});
        pie.setSeriesData(new SeriesRenderData("countries", new Integer[]{17098242, 9984670, 9826675, 9596961}));
        Map<String, Object> map = new HashMap<>();
        map.put("pieChart", pie);
        // 编译模板、渲染数据
        XWPFTemplate template = XWPFTemplate.compile("src/main/resources/charpter3/template3.docx").render(map);
        FileOutputStream out = new FileOutputStream("src/main/resources/charpter3/template3_out.docx");
        // 输出到流
        template.write(out);
        out.flush();
        out.close();
        template.close();
    }
}
