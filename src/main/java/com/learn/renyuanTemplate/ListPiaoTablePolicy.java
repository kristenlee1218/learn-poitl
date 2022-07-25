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

import java.util.ArrayList;
import java.util.List;

/**
 * @author ：Kristen
 * @date ：2022/7/25
 * @description :
 */
public class ListPiaoTablePolicy extends AbstractRenderPolicy<Object> {

    String[] voteType = new String[]{"A1", "A2", "A3", "B", "C"};
    public static String[][] data = new String[][]{{"刘备", "董事长", "80.00", "1", "78.96", "1", "74.98", "1", "80.56", "1", "89.22", "1", "78.25", "1"}, {"关羽", "总经理", "80.00", "2", "78.96", "2", "74.98", "2", "80.56", "2", "89.22", "2", "78.25", "2"}};

    // 计算行和列
    int col;
    int row;

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
        col = (voteType.length + 1) * 2 + 3;
        row = data.length + 3;

        // 当前位置插入表格
        XWPFTable table = bodyContainer.insertNewTable(run, row, col);
        this.setTableStyle(table);
        this.setTableTitle(table);
        this.setTableHeader(table);
    }

    // 整个 table 的样式在此设置
    public void setTableStyle(XWPFTable table) {
        // 设置 A4 幅面的平铺类型和列数
        TableTools.widthTable(table, MiniTableRenderData.WIDTH_A4_NARROW_FULL, col);

        // 设置 border
        TableTools.borderTable(table, 10);
        for (XWPFTableRow tableRow : table.getRows()) {
            for (int i = 0; i < tableRow.getTableCells().size(); i++) {
                tableRow.getCell(i).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
                tableRow.getCell(i).setWidth("500");
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
        TableStyle tableStyle = new TableStyle();
        tableStyle.setAlign(STJc.CENTER);
        tableStyle.setBackgroundColor("DCDCDC");
        String title = "{{title}}";
        RowRenderData header0 = RowRenderData.build(new TextRenderData(title, cellStyle));
        header0.setRowStyle(tableStyle);
        TableTools.mergeCellsHorizonal(table, 0, 0, col - 1);
        MiniTableRenderPolicy.Helper.renderRow(table, 0, header0);
    }

    public void setTableHeader(XWPFTable table) {
        // 构建第二行的数组、并垂直合并与水平合并
        String[] strHeader1 = new String[3 + voteType.length + 1];
        strHeader1[0] = "序号";
        strHeader1[1] = "姓名";
        strHeader1[2] = "职务";
        strHeader1[strHeader1.length - 1] = "全体";

        int start = 3;
        // 水平合并 group 的名字
        for (int i = 0; i < voteType.length; i++) {
            TableTools.mergeCellsHorizonal(table, 1, start, start++);
            start++;
        }

        for (int i = 0; i < voteType.length; i++) {
            strHeader1[i + 3] = voteType[i];
            //TableTools.mergeCellsVertically(table, i, 1, 2);
        }

        // 构建第三行的数组
//        String[] strHeader2 = new String[col];
//        int index = 5;
//        for (String[] str : items) {
//            strHeader2[index] = "小计";
//            for (String s : str) {
//                strHeader2[++index] = s;
//            }
//            index++;
//        }

        // 构建第二、三行
        Style style = this.getCellStyle();
        RowRenderData header1 = this.build(style, strHeader1);
        // RowRenderData header2 = this.build(style, strHeader2);
        TableStyle tableStyle = this.getTableStyle();
        header1.setRowStyle(tableStyle);
        // header2.setRowStyle(tableStyle);
        MiniTableRenderPolicy.Helper.renderRow(table, 1, header1);
        //MiniTableRenderPolicy.Helper.renderRow(table, 2, header2);
    }

    // 根据 String[] 构建一行的数据，同一行使用一个 Style
    public RowRenderData build(Style style, String... cellStr) {
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
        cellStyle.setFontFamily("宋体");
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

    // 设置 cell 格样式
    public Style getDataCellStyle() {
        Style cellStyle = new Style();
        cellStyle.setFontFamily("宋体");
        cellStyle.setFontSize(8);
        cellStyle.setColor("000000");
        return cellStyle;
    }
}
