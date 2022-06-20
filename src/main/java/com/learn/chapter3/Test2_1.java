package com.learn.chapter3;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.data.ChartMultiSeriesRenderData;
import com.deepoove.poi.data.SeriesRenderData;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author ：Kristen
 * @date ：2022/6/7
 * @description : poi-tl 操作多图表
 */
public class Test2_1 {
    public static void main(String[] args) throws IOException {
        ChartMultiSeriesRenderData chart = new ChartMultiSeriesRenderData();
        chart.setChartTitle("MyChart");
        chart.setCategories(new String[]{"中文", "English"});
        List<SeriesRenderData> seriesRenderData = new ArrayList<>();
        seriesRenderData.add(new SeriesRenderData("countries", new Double[]{15.0, 6.0}));
        seriesRenderData.add(new SeriesRenderData("speakers", new Double[]{223.0, 119.0}));
        Map<String, Object> map = new HashMap<>();
        chart.setSeriesDatas(seriesRenderData);
        map.put("barChart", chart);
        // 编译模板、渲染数据
        XWPFTemplate template = XWPFTemplate.compile("D:\\test-poitl\\template.docx").render(map);
        FileOutputStream out = new FileOutputStream("D:\\test-poitl\\output.docx");
        // 输出到流
        template.write(out);
        out.flush();
        out.close();
        template.close();
    }
}
