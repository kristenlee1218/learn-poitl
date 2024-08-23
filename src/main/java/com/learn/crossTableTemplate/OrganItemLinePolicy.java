package com.learn.crossTableTemplate;

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
 * @date ：2024/8/23
 * @description :
 */
public class OrganItemLinePolicy {
    public static void main(String[] args) throws IOException {
        // 编译模板、渲染数据
        XWPFTemplate template = XWPFTemplate.compile("D:\\test-poitl\\ele02.docx").render(new HashMap<String, Object>() {{
            ChartMultiSeriesRenderData chart = new ChartMultiSeriesRenderData();
            chart.setChartTitle("各测评主体评价情况对比");
            chart.setCategories(new String[]{"政治忠诚", "政治担当", "社会责任", "改革创新", "管理效能", "风险管控", "选人用人", "基层党建", "团结协作", "联系群众"});
            List<SeriesRenderData> seriesRenderData = new ArrayList<>();
            seriesRenderData.add(new SeriesRenderData("A1票", new Double[]{100.00, 99.00, 91.00, 97.00, 87.00, 100.00, 99.00, 78.00, 87.00, 96.00}));
            seriesRenderData.add(new SeriesRenderData("A2票", new Double[]{96.00, 89.00, 96.00, 92.00, 100.00, 86.00, 97.00, 78.00, 82.00, 100.00,}));
            seriesRenderData.add(new SeriesRenderData("A3票", new Double[]{91.00, 78.00, 98.00, 96.00, 93.00, 85.00, 97.00, 68.00, 81.00, 100.00,}));
            seriesRenderData.add(new SeriesRenderData("A4票", new Double[]{92.00, 87.00, 98.00, 99.00, 94.00, 84.00, 96.00, 58.00, 80.00, 80.00,}));
            seriesRenderData.add(new SeriesRenderData("B票", new Double[]{93.00, 95.00, 98.00, 96.00, 95.00, 96.00, 83.00, 88.00, 99.00, 82.00,}));
            seriesRenderData.add(new SeriesRenderData("C票", new Double[]{94.00, 100.00, 98.00, 96.00, 100.00, 96.00, 82.00, 88.00, 99.00, 79.00,}));
            chart.setSeriesDatas(seriesRenderData);
            put("lineChart", chart);
        }});
        FileOutputStream out = new FileOutputStream("D:\\test-poitl\\ele02_out.docx");
        // 输出到流
        template.write(out);
        out.flush();
        out.close();
        template.close();
    }
}
