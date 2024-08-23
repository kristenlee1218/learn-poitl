package com.learn.leaderTemplate;

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
 * @date ：2024/8/8
 * @description :
 */
public class LeaderItemLinePolicy {
    public static void main(String[] args) throws IOException {
        // 编译模板、渲染数据
        XWPFTemplate template = XWPFTemplate.compile("D:\\test-poitl\\ele02.docx").render(new HashMap<String, Object>() {{
            ChartMultiSeriesRenderData chart = new ChartMultiSeriesRenderData();
            chart.setChartTitle("各测评主体评价情况对比\n");
            chart.setCategories(new String[]{"政治品质", "政治本领", "创新精神", "创新成果", "经营管理能力", "抓党建强党建能力", "担当作为", "履职绩效", "一岗双责", "廉洁从业"});
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
