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

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;

/**
 * @author ：Kristen
 * @date ：2023/8/31
 * @description :
 */
public class SelectPeople1DuibiTablePolicy extends AbstractRenderPolicy<Object> {

    public static String option = "4:不了解:0;6:不好:0;8:一般:0;10:好:0";
    public static String[] question = new String[]{"1、对本单位选人用人工作的总体评价", "2、对本单位从严管理监督干部情况的评价"};
    public static String[] itemId = new String[]{"organ21", "organ22"};
    public int dataSize = 1;

    // 计算行和列
    int col;
    int row;
    LinkedHashMap<String, Integer> optionMap;
    int colBase = 2;
    int rowBase = 3;

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
        optionMap = this.splitOption(option);
        optionMap.put("好与一般的比例", 0);
        row = rowBase + dataSize;
        col = optionMap.size() * 2 + colBase + 1;

        // 当前位置插入表格
        XWPFTable table = bodyContainer.insertNewTable(run, row, col);
        this.setTableStyle(table);
        this.setTableTitle(table);
        this.setTableHeader(table);
        this.setTableTag(table);
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
                if (i == 0) {
                    tableRow.getCell(i).setWidth("2000");
                    table.setTableAlignment(TableRowAlign.LEFT);
                    table.setCellMargins(2, 0, 2, 0);
                } else {
                    tableRow.getCell(i).setWidth("500");
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
        strHeader1[strHeader1.length - 1] = "两项好和一般比率平均值";

        // 第1行垂直合并单元格
        TableTools.mergeCellsVertically(table, 0, 1, 2);
        TableTools.mergeCellsVertically(table, 1, 1, 2);
        TableTools.mergeCellsVertically(table, col - 1, 1, 2);

        int header1Start = colBase;
        int header1End;
        for (int i = 0; i < question.length; i++) {
            header1End = optionMap.size() + header1Start - 1;
            TableTools.mergeCellsHorizonal(table, 1, header1Start, header1End);
            header1Start++;
        }

        // 第2行值的数组
        String[] strHeader2 = new String[optionMap.size() * question.length + colBase + 1];
        int index = colBase;
        for (int i = 0; i < question.length; i++) {
            for (int j = 0; j < optionMap.keySet().toArray().length; j++) {
                strHeader2[index++] = optionMap.keySet().toArray()[j].toString();
            }
        }

        // 构建第1-2行
        Style style = this.getCellStyle();
        RowRenderData header1 = this.build(strHeader1, style);
        RowRenderData header2 = this.build(strHeader2, style);
        TableStyle tableStyle = this.getTableStyle();
        header1.setRowStyle(tableStyle);
        header2.setRowStyle(tableStyle);
        MiniTableRenderPolicy.Helper.renderRow(table, 1, header1);
        MiniTableRenderPolicy.Helper.renderRow(table, 2, header2);
    }

    public void setTableTag(XWPFTable table) {
        for (int i = 0; i < dataSize; i++) {
            // 设置 tag（票种类型部分）
            String[] str = new String[col];
            str[0] = "{{sequence_" + i + "}}";
            str[1] = "{{organshortname_" + i + "}}";

            // 构建
            Style style = this.getCellStyle();
            RowRenderData tag = this.build(str, style);
            TableStyle tableStyle = this.getTableStyle();
            tag.setRowStyle(tableStyle);
            MiniTableRenderPolicy.Helper.renderRow(table, i + rowBase, tag);
        }
    }

    // 将题目的选项 如：“4:不了解:0;6:不好:0;8:一般:0;10:好:0” 存入 map，key 为显示的值，value 为分值
    public LinkedHashMap<String, Integer> splitOption(String option) {
        LinkedHashMap<String, Integer> optionMap = new LinkedHashMap<>();
        String[] strArray = option.replaceAll(":0", "").split(";");
        for (int i = strArray.length - 1; i >= 0; i--) {
            String[] str = strArray[i].split(":");
            optionMap.put(str[1], Integer.valueOf(str[0]));
        }
        return optionMap;
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
