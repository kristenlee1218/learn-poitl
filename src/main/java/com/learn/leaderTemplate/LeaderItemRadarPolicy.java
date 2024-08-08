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
public class LeaderItemRadarPolicy {
    public static void main(String[] args) throws IOException {
        // 编译模板、渲染数据
        XWPFTemplate template = XWPFTemplate.compile("D:\\test-poitl\\ele02.docx").render(
                new HashMap<String, Object>() {{
                    ChartMultiSeriesRenderData chart = new ChartMultiSeriesRenderData();
                    chart.setChartTitle("各测评指标与领导人员平均分对比情况");
                    chart.setCategories(new String[]{"政治品质", "政治本领", "创新精神", "创新成果", "经营管理能力", "抓党建强党建能力", "担当作为", "履职绩效", "一岗双责", "廉洁从业"});
                    List<SeriesRenderData> seriesRenderData = new ArrayList<>();
                    seriesRenderData.add(new SeriesRenderData("张三", new Double[]{100.00, 99.00, 98.00, 97.00, 96.00, 100.00, 99.00, 98.00, 97.00, 96.00}));
                    seriesRenderData.add(new SeriesRenderData("领导人员平均", new Double[]{96.00, 97.00, 98.00, 99.00, 100.00, 96.00, 97.00, 98.00, 99.00, 100.00,}));
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
