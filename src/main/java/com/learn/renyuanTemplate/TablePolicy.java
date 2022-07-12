package com.learn.renyuanTemplate;

import com.deepoove.poi.data.MiniTableRenderData;
import com.deepoove.poi.policy.AbstractRenderPolicy;
import com.deepoove.poi.render.RenderContext;
import com.deepoove.poi.util.TableTools;
import com.deepoove.poi.xwpf.BodyContainer;
import com.deepoove.poi.xwpf.BodyContainerFactory;
import org.apache.poi.xwpf.usermodel.*;

/**
 * @author ：Kristen
 * @date ：2022/7/12
 * @description :
 */
public class TablePolicy extends AbstractRenderPolicy<Object> {

    //第一种测试情况（信息中心_2020年第1批综合测评统计报表(1655863588387)）(14)
    public static String[] group = new String[]{"对党忠诚", "勇于创新", "治企有方", "兴企有为", "清正廉洁"};
    public static String[][] item = new String[][]{{"政治品质", "政治本领"}, {"创新精神", "创新成果"}, {"经营管理能力", "抓党建强党建能力"}, {"担当作为", "履职绩效"}, {"一岗双责", "廉洁从业"}};
    public static String[] people = new String[]{"刘备", "诸葛亮", "关羽", "张飞", "赵云", "黄忠", "马超"};
    public static int year = 2022;
    public static String depart = "信息中心";

    // 计算行和列
    int col = countCol(item) + 5;
    int row = people.length + 3;

    @Override
    protected void afterRender(RenderContext<Object> renderContext) {
        // 清空标签
        clearPlaceholder(renderContext, true);
    }

    @Override
    public void doRender(RenderContext<Object> renderContext) {
        XWPFRun run = renderContext.getRun();
        // 当前位置的容器
        BodyContainer bodyContainer = BodyContainerFactory.getBodyContainer(run);

        // 当前位置插入表格
        XWPFTable table = bodyContainer.insertNewTable(run, row, col);
        setTableStyle(table);
    }

    // 整个 table 的样式在此设置
    public void setTableStyle(XWPFTable table) {
        // 设置A4幅面的平铺类型和列数
        TableTools.widthTable(table, MiniTableRenderData.WIDTH_A4_FULL, col);
        // 设置 border
        TableTools.borderTable(table, 9);
        for (XWPFTableRow tableRow : table.getRows()) {
            for (int i = 0; i < tableRow.getTableCells().size(); i++) {
                tableRow.getCell(i).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
            }
        }
        table.setCellMargins(2, 2, 2, 2);
        table.setTableAlignment(TableRowAlign.CENTER);
    }

    // 计算所有分组的项的个数
    public int countCol(String[][] str) {
        int count = 0;
        int index = 0;
        for (String[] s : str) {
            index++;
            for (int j = 0; j < s.length; j++) {
                count++;
            }
        }
        return count + index;
    }
}
