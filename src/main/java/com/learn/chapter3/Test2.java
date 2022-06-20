package com.learn.chapter3;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.data.ChartMultiSeriesRenderData;
import com.deepoove.poi.data.SeriesRenderData;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

/**
 * @author ：Kristen
 * @date ：2022/6/7
 * @description : poi-tl 操作多图表
 */
public class Test2 {
    public static void main(String[] args) throws IOException {
        // 编译模板、渲染数据
        XWPFTemplate template = XWPFTemplate.compile("src/main/resources/charpter3/template2.docx").render(
                new HashMap<String, Object>() {{
                    ChartMultiSeriesRenderData chart = new ChartMultiSeriesRenderData();
                    chart.setChartTitle("MyChart");
                    chart.setCategories(new String[]{"中文", "English"});
                    List<SeriesRenderData> seriesRenderData = new ArrayList<>();
                    seriesRenderData.add(new SeriesRenderData("countries", new Double[]{15.0, 6.0}));
                    seriesRenderData.add(new SeriesRenderData("speakers", new Double[]{223.0, 119.0}));
                    chart.setSeriesDatas(seriesRenderData);
                    put("barChart", chart);
                }});
        FileOutputStream out = new FileOutputStream("src/main/resources/charpter3/template2_out.docx");
        // 输出到流
        template.write(out);
        out.flush();
        out.close();
        template.close();
    }
}
