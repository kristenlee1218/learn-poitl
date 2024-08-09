package com.learn.leaderTemplate;

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

import java.util.ArrayList;
import java.util.List;

/**
 * @author ：Kristen
 * @date ：2024/8/7
 * @description : 反馈报告中的均分对比表
 */
public class LeaderItemAverageDuibiPolicy extends AbstractRenderPolicy<Object> {

    int col = item.length + 2;
    int row = 3;
    public static String[] item = new String[]{"政治品质", "政治本领", "创新精神", "创新成果", "经营管理能力", "抓党建强党建能力", "担当作为", "履职绩效", "一岗双责", "廉洁从业"};
    public static String[] itemId = new String[]{"leader01", "leader02", "leader03", "leader04", "leader05", "leader06", "leader07", "leader08", "leader09", "leader10"};


    @Override
    public void afterRender(RenderContext<Object> renderContext) {
        // 清空标签
        clearPlaceholder(renderContext, true);
    }

    @Override
    public void doRender(RenderContext<Object> renderContext) throws Exception {
        XWPFRun run = renderContext.getRun();
        // 当前位置的容器
        BodyContainer bodyContainer = BodyContainerFactory.getBodyContainer(run);
        // 计算行列

        // 当前位置插入表格
        XWPFTable table = bodyContainer.insertNewTable(run, row, col);
        // 当前位置插入表格
        this.setTableStyle(table);
        this.setItem(table);
        this.setCellTag(table);
    }

    // 整个 table 的样式在此设置
    public void setTableStyle(XWPFTable table) {
        // 设置 A4 幅面的平铺类型和列数
        TableTools.widthTable(table, 27, col);

        // 设置 border
        TableTools.borderTable(table, 10);
        for (XWPFTableRow tableRow : table.getRows()) {
            for (int i = 0; i < tableRow.getTableCells().size(); i++) {
                tableRow.getCell(i).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
                tableRow.getCell(i).setWidth("20");
            }
            tableRow.setHeight(500);
        }
        table.setCellMargins(2, 2, 2, 2);
        table.setTableAlignment(TableRowAlign.CENTER);
    }

    public void setItem(XWPFTable table) {
        String[] strHeader0 = new String[col];
        System.arraycopy(item, 0, strHeader0, 1, item.length);
        strHeader0[strHeader0.length - 1] = "总分";
        Style style = this.getCellStyle();
        RowRenderData header1 = this.build(strHeader0, style);
        TableStyle tableStyle = this.getTableStyle();
        header1.setRowStyle(tableStyle);
        MiniTableRenderPolicy.Helper.renderRow(table, 0, header1);
    }

    public void setCellTag(XWPFTable table) {
        // 数组填充值
        String[] strHeader1 = new String[col];
        String[] strHeader2 = new String[col];
        for (int i = 0; i < itemId.length; i++) {
            strHeader1[i + 1] = "avg#" + itemId[i];
        }
        strHeader1[0] = "指标得分";
        strHeader1[strHeader1.length - 1] = "avg";
        for (int i = 0; i < itemId.length; i++) {
            strHeader2[i + 1] = "listavg#" + itemId[i];
        }
        strHeader1[0] = "平均分";
        strHeader1[strHeader1.length - 1] = "listavg";

        // 设置样式与构建行对象
        Style style = this.getCellStyle();
        RowRenderData header1 = this.build(strHeader1, style);
        RowRenderData header2 = this.build(strHeader2, style);
        TableStyle tableStyle = this.getTableStyle();
        header1.setRowStyle(tableStyle);
        header2.setRowStyle(tableStyle);
        MiniTableRenderPolicy.Helper.renderRow(table, 1, header1);
        MiniTableRenderPolicy.Helper.renderRow(table, 2, header2);

    }

    // 根据 String[] 构建一行的数据，同一行使用一个 Style
    public RowRenderData build(String[] cellStr, Style style) {
        List<TextRenderData> data = new ArrayList<>();
        if (null != cellStr) {
            for (String s : cellStr) {
                data.add(new TextRenderData(s, style));
            }
        }
        return new RowRenderData(data, null);
    }

    // 设置 cell 格样式
    public Style getCellStyle() {
        Style cellStyle = new Style();
        cellStyle.setFontFamily("仿宋");
        cellStyle.setFontSize(10);
        cellStyle.setColor("000000");
        return cellStyle;
    }

    // 设置 table 格样式
    public TableStyle getTableStyle() {
        TableStyle tableStyle = new TableStyle();
        tableStyle.setAlign(STJc.CENTER);
        return tableStyle;
    }
}
