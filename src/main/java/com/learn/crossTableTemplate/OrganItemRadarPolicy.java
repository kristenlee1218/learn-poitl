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
public class OrganItemRadarPolicy {
    public static void main(String[] args) throws IOException {
        // 编译模板、渲染数据
        XWPFTemplate template = XWPFTemplate.compile("D:\\test-poitl\\ele02.docx").render(new HashMap<String, Object>() {{
            ChartMultiSeriesRenderData chart = new ChartMultiSeriesRenderData();
            chart.setChartTitle("各测评指标情况");
            chart.setCategories(new String[]{"政治忠诚", "政治担当", "社会责任", "改革创新", "管理效能", "风险管控", "选人用人", "基层党建", "团结协作", "联系群众"});
            List<SeriesRenderData> seriesRenderData = new ArrayList<>();
            seriesRenderData.add(new SeriesRenderData("单位1", new Double[]{96.00, 97.00, 98.00, 99.00, 100.00, 96.00, 97.00, 98.00, 99.00, 100.00,}));
            chart.setSeriesDatas(seriesRenderData);
            put("barChart", chart);
        }});
        FileOutputStream out = new FileOutputStream("D:\\test-poitl\\ele02_out.docx");
        // 输出到流
        template.write(out);
        out.flush();
        out.close();
        template.close();
    }
}
