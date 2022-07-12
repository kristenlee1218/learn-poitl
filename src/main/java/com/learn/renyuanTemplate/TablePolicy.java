package com.learn.renyuanTemplate;

import com.deepoove.poi.data.MiniTableRenderData;
import com.deepoove.poi.data.RowRenderData;
import com.deepoove.poi.data.TextRenderData;
import com.deepoove.poi.data.style.Style;
import com.deepoove.poi.data.style.TableStyle;
import com.deepoove.poi.policy.AbstractRenderPolicy;
import com.deepoove.poi.policy.MiniTableRenderPolicy;
import com.deepoove.poi.render.RenderContext;
import com.deepoove.poi.util.TableTools;
import com.deepoove.poi.xwpf.BodyContainer;
import com.deepoove.poi.xwpf.BodyContainerFactory;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJc;

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
        setTableTitle(table);
        setTableHeader(table);
    }

    // 整个 table 的样式在此设置
    public void setTableStyle(XWPFTable table) {
        // 设置A4幅面的平铺类型和列数
        TableTools.widthTable(table, MiniTableRenderData.WIDTH_A4_NARROW_FULL, col);
        // 设置 border
        TableTools.borderTable(table, 9);
        for (XWPFTableRow tableRow : table.getRows()) {
            for (int i = 0; i < tableRow.getTableCells().size(); i++) {
                tableRow.getCell(i).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
                tableRow.getCell(i).setWidth("1000");
            }
        }
        table.setCellMargins(2, 2, 2, 2);
        table.setTableAlignment(TableRowAlign.CENTER);
    }

    // 设置第一行标题的样式
    public void setTableTitle(XWPFTable table) {
        Style cellStyle = new Style();
        cellStyle.setFontSize(12);
        cellStyle.setColor("000000");
        cellStyle.setFontFamily("黑体");
        TableStyle style = new TableStyle();
        style.setAlign(STJc.CENTER);
        style.setBackgroundColor("DCDCDC");
        String title = depart + "领导人员" + year + "年度综合测评汇总表";
        RowRenderData header0 = RowRenderData.build(new TextRenderData(title, cellStyle));
        header0.setRowStyle(style);
        TableTools.mergeCellsHorizonal(table, 0, 0, col - 1);
        MiniTableRenderPolicy.Helper.renderRow(table, 0, header0);
    }

    // 设置第二、三行标题的样式
    public void setTableHeader(XWPFTable table) {
        String[] strHeader1 = new String[5 + group.length];
        strHeader1[0] = "序号";
        strHeader1[1] = "姓名";
        strHeader1[2] = "职务";
        strHeader1[3] = "汇总得分";
        strHeader1[4] = "排名";

        int start = 5;
        int end;
        // 水平合并 group 的名字
        for (String[] value : item) {
            end = (value.length + 1);
            TableTools.mergeCellsHorizonal(table, 1, start, start + end - 1);
            start++;
        }

        // 构建第二行的数组、并垂直合并
        for (int i = 0; i < group.length; i++) {
            strHeader1[i + 5] = group[i];
            TableTools.mergeCellsVertically(table, i, 1, 2);
        }

        String[] strHeader2 = new String[col];
        int index = 5;
        for (String[] strings : item) {
            strHeader2[index] = "小计";
            for (String string : strings) {
                strHeader2[++index] = string;
            }
            ++index;
        }

        // 构建第二行
        RowRenderData header1 = RowRenderData.build(strHeader1);
        RowRenderData header2 = RowRenderData.build(strHeader2);
        TableStyle style = new TableStyle();
        style.setAlign(STJc.CENTER);
        header1.setRowStyle(style);
        header2.setRowStyle(style);
        MiniTableRenderPolicy.Helper.renderRow(table, 1, header1);
        MiniTableRenderPolicy.Helper.renderRow(table, 2, header2);
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
