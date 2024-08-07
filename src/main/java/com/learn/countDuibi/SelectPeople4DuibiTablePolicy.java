package com.learn.countDuibi;

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

import java.util.*;

/**
 * @author ：Kristen
 * @date ：2023/8/31
 * @description :
 */
public class SelectPeople4DuibiTablePolicy extends AbstractRenderPolicy<Object> {

    public static String[] question = new String[]{"1、对本单位选人用人工作的总体评价", "2、对本单位从严管理监督干部情况的评价"};
    public static String[] itemId = new String[]{"organ21", "organ22"};
    public static String[][] data = new String[][]{{"1", "单位名称1", "100%", "99%", "99.55%"}, {"2", "单位名称2", "100%", "99%", "99.55%"}, {"3", "单位名称3", "100%", "99%", "99.55%"}, {"4", "单位名称4", "100%", "99%", "99.55%"}};
    public int dataSize = 4;

    // 计算行和列
    int col;
    int row;
    int colBase = 2;
    int rowBase = 2;

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
        row = rowBase + dataSize;
        col = question.length + colBase + 1;

        // 当前位置插入表格
        XWPFTable table = bodyContainer.insertNewTable(run, row, col);
        this.setTableStyle(table);
        this.setTableTitle(table);
        this.setTableHeader(table);
//        this.setTableTag(table);
        this.setTableData(table);
    }

    // 整个 table 的样式在此设置
    public void setTableStyle(XWPFTable table) {
        // 设置 A4 幅面的平铺类型和列数
        TableTools.widthTable(table, 2100, col);

        // 设置 border
        TableTools.borderTable(table, 10);
        for (XWPFTableRow tableRow : table.getRows()) {
            for (int i = 0; i < tableRow.getTableCells().size(); i++) {
                tableRow.getCell(i).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
                if (i == 0) {
                    tableRow.getCell(i).setWidth("10%");
                    table.setTableAlignment(TableRowAlign.LEFT);
                    table.setCellMargins(2, 0, 2, 0);
                } else {
                    tableRow.getCell(i).setWidth("20%");
                    table.setTableAlignment(TableRowAlign.CENTER);
                    table.setCellMargins(2, 2, 2, 2);
                }
            }
        }
    }

    // 设置第一行标题的样式
    public void setTableTitle(XWPFTable table) {
        Style cellStyle = new Style();
        cellStyle.setFontSize(12);
        cellStyle.setColor("000000");
        cellStyle.setFontFamily("黑体");
        TableStyle tableStyle = new TableStyle();
        tableStyle.setAlign(STJc.CENTER);
        tableStyle.setBackgroundColor("F5F5F5");
        String title = "{{title}}";
        RowRenderData header0 = RowRenderData.build(new TextRenderData(title, cellStyle));
        header0.setRowStyle(tableStyle);
        TableTools.mergeCellsHorizonal(table, 0, 0, col - 1);
        MiniTableRenderPolicy.Helper.renderRow(table, 0, header0);
    }

    public void setTableHeader(XWPFTable table) {
        // 第1行值的数组
        String[] strHeader1 = new String[colBase + question.length + 1];
        strHeader1[0] = "序号";
        strHeader1[1] = "单位名称";
        System.arraycopy(question, 0, strHeader1, colBase, question.length);
        strHeader1[strHeader1.length - 1] = "平均好与一般的比例";

        // 构建第1-2行
        Style style = this.getCellStyle();
        RowRenderData header1 = this.build(strHeader1, style);
        TableStyle tableStyle = this.getTableStyle();
        header1.setRowStyle(tableStyle);
        MiniTableRenderPolicy.Helper.renderRow(table, 1, header1);
    }

    public void setTableTag(XWPFTable table) {
        for (int i = 0; i < dataSize; i++) {
            // 设置 tag（票种类型部分）
            String[] str = new String[col];
            str[0] = "{{sort_" + (colBase + question.length) + "_" + i + "}}";
            str[1] = "{{organshortname_" + i + "}}";
            for (int j = 0; j < itemId.length; j++) {
                str[colBase + j] = "{{rate#" + itemId[j] + ">7_" + i + "#}}";
            }
            str[str.length - 1] = "{{average(2,3)_" + i + "_###sort}}";

            // 构建
            Style style = this.getCellStyle();
            RowRenderData tag = this.build(str, style);
            TableStyle tableStyle = this.getTableStyle();
            tag.setRowStyle(tableStyle);
            MiniTableRenderPolicy.Helper.renderRow(table, i + rowBase, tag);
        }
    }

    public void setTableData(XWPFTable table) {
        for (int i = 0; i < data.length; i++) {
            // 设置 tag（票种类型部分）
            String[] str = new String[col];
            System.arraycopy(data[i], 0, str, 0, data[i].length);

            // 构建
            Style style = this.getCellStyle();
            RowRenderData tag = this.build(str, style);
            TableStyle tableStyle = this.getTableStyle();
            tag.setRowStyle(tableStyle);
            MiniTableRenderPolicy.Helper.renderRow(table, i + rowBase, tag);
        }
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

    // 设置标题 cell 格样式
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

    // 设置数据 cell 格样式
    public Style getDataCellStyle() {
        Style cellStyle = new Style();
        cellStyle.setFontFamily("宋体");
        cellStyle.setFontSize(8);
        cellStyle.setColor("000000");
        return cellStyle;
    }
}
