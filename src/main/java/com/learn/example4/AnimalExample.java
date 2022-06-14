package com.learn.example4;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import com.deepoove.poi.data.ChartMultiSeriesRenderData;
import com.deepoove.poi.data.PictureRenderData;
import com.deepoove.poi.data.SeriesRenderData;
import com.deepoove.poi.policy.reference.MultiSeriesChartTemplateRenderPolicy;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author ：Kristen
 * @date ：2022/6/14
 * @description :
 */
@DisplayName("AnimalExmample test case")
public class AnimalExample {

    @Test
    public void testChart() throws Exception {

        Configure configure = Configure.newBuilder().bind("chart", new MultiSeriesChartTemplateRenderPolicy()).build();
        HashMap<String, Object> data1 = new HashMap<String, Object>() {
            {
                put("name", "大象");
                ChartMultiSeriesRenderData chart = new ChartMultiSeriesRenderData();
                chart.setChartTitle("大象生存现状");
                chart.setCategories(new String[]{"2018年", "2019年", "2020年"});
                List<SeriesRenderData> seriesRenderData = new ArrayList<SeriesRenderData>();
                SeriesRenderData s0 = new SeriesRenderData();
                s0.setName("成年象");
                s0.setValues(new Integer[]{500, 600, 700});
                seriesRenderData.add(s0);
                SeriesRenderData s1 = new SeriesRenderData();
                s1.setName("幼象");
                s1.setValues(new Integer[]{200, 300, 400});
                seriesRenderData.add(s1);
                SeriesRenderData s2 = new SeriesRenderData();
                s2.setName("全部");
                s2.setValues(new Integer[]{700, 900, 1100});
                seriesRenderData.add(s2);
                chart.setSeriesDatas(seriesRenderData);
                put("chart", chart);
            }
        };
        HashMap<String, Object> data2 = new HashMap<String, Object>() {
            {
                put("name", "长颈鹿");
                put("picture", new PictureRenderData(100, 120, ".png", Files.newInputStream(Paths.get("D:\\test-poitl\\lu.png"))));
                ChartMultiSeriesRenderData chart = new ChartMultiSeriesRenderData();
                chart.setChartTitle("长颈鹿生存现状");
                chart.setCategories(new String[]{"2018年", "2019年", "2020年"});
                List<SeriesRenderData> seriesRenderData = new ArrayList<>();
                SeriesRenderData s0 = new SeriesRenderData();
                s0.setName("成年鹿");
                s0.setValues(new Integer[]{500, 600, 700});
                seriesRenderData.add(s0);
                SeriesRenderData s1 = new SeriesRenderData();
                s1.setName("幼鹿");
                s1.setValues(new Integer[]{200, 300, 400});
                seriesRenderData.add(s1);
                chart.setSeriesDatas(seriesRenderData);
                put("chart", chart);

            }
        };
        List<Map<String, Object>> animals = new ArrayList<>();
        animals.add(data1);
        animals.add(data2);
        Map<String, Object> data = new HashMap<String, Object>() {
            {
                put("animals", animals);
            }
        };
        XWPFTemplate template = XWPFTemplate.compile("D:\\test-poitl\\template4.docx", configure).render(data);
        template.writeToFile("D:\\test-poitl\\template4_out.docx");
    }
}
