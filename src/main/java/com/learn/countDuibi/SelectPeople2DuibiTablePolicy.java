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
public class SelectPeople2DuibiTablePolicy extends AbstractRenderPolicy<Object> {

    public static String[] voteName = new String[]{"班子成员占比", "原领导班子和中层干部占比", "职工代表占比", "全体人员占比"};
    public static String[] voteType = new String[]{"A1/A2/A3", "B", "C", ""};

    public static String[] question = new String[]{"3、您认为本单位选人用人工作存在的主要问题是什么？（可多选）"};
    public static String option = "1:落实党中央关于领导班子和干部队伍建设工作要求有差距:0;2:选人用人把关不严、质量不高:0;3:坚持事业为上不够，不能做到以事择人、人岗相适:0;4:激励担当作为用人导向不鲜明，论资排辈情况严重:0;5:选人用人“个人说了算”:0";
    public static String[] data = new String[]{"1", "2", "3", "4", "5", "6", "7", "8", "9", "10"};
    public static String[] itemId = new String[]{"organ23"};
    public int dataSize = 10;

    int col;
    int row;
    LinkedHashMap<String, Integer> optionMap;
    int colBase = 2;
    int rowBase = 4;

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
        // 计算行列
        row = rowBase + dataSize;
        col = optionMap.size() * voteName.length + colBase;

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

    // 设置 header
    public void setTableHeader(XWPFTable table) {

        // 第1行值的数组
        String[] strHeader1 = new String[colBase + question.length];
        strHeader1[0] = "序号";
        strHeader1[1] = "单位名称";
        System.arraycopy(question, 0, strHeader1, colBase, question.length);

        // 第1行垂直合并单元格
        for (int i = 0; i < colBase; i++) {
            TableTools.mergeCellsVertically(table, i, 1, 3);
        }

        // 第1行水平合并单元格
        TableTools.mergeCellsHorizonal(table, 1, colBase, col - 1);

        // 第2行值的数组
        String[] strHeader2 = new String[colBase + optionMap.size()];
        Object[] optionStr = optionMap.keySet().toArray();
        for (int i = 0; i < optionStr.length; i++) {
            strHeader2[i + colBase] = "（" + (i + 1) + "）、" + optionStr[i].toString();
        }

        // 第2行水平合并单元格
        int start = colBase;
        int end;
        for (int i = 0; i < optionMap.size(); i++) {
            end = start + voteName.length;
            TableTools.mergeCellsHorizonal(table, 2, start, end - 1);
            start++;
        }

        // 第3行值的数组
        String[] strHeader3 = new String[col];
        int index = colBase;
        for (int i = 0; i < optionMap.size(); i++) {
            for (String s : voteName) {
                strHeader3[index++] = s;
            }
        }

        // 构建第1-3行
        Style style = this.getCellStyle();
        RowRenderData header1 = this.build(strHeader1, style);
        RowRenderData header2 = this.build(strHeader2, style);
        RowRenderData header3 = this.build(strHeader3, style);
        TableStyle tableStyle = this.getTableStyle();
        header1.setRowStyle(tableStyle);
        header2.setRowStyle(tableStyle);
        header3.setRowStyle(tableStyle);
        MiniTableRenderPolicy.Helper.renderRow(table, 1, header1);
        MiniTableRenderPolicy.Helper.renderRow(table, 2, header2);
        MiniTableRenderPolicy.Helper.renderRow(table, 3, header3);
    }

    //sectrate#REPLACE(CONCAT(organ23,''),'10','')  like '%1%'#@A1;A2;A3
    public void setTableTag(XWPFTable table) {
        for (int i = 0; i < dataSize; i++) {
            String[] tag = new String[col];
            tag[0] = "{{sequence}}";
            tag[1] = "{{organshortname}}";
            int index = colBase;
            Object[] optionValue = optionMap.values().toArray();
            for (int j = 0; j < optionValue.length; j++) {
                for (String s : voteType) {
                    tag[index++] = "{{sectrate#REPLACE(CONCAT(" + itemId[0] + ",''),'" + optionValue[j] + "','')  like '%" + j + "%'#@" + s.replace("/", "_") + "_" + i + "}}";
                }
            }
            Style style = this.getCellStyle();
            RowRenderData tagData = this.build(tag, style);
            TableStyle tableStyle = this.getTableStyle();
            tagData.setRowStyle(tableStyle);
            MiniTableRenderPolicy.Helper.renderRow(table, rowBase + i, tagData);
        }
    }

    public LinkedHashMap<String, Integer> splitOption(String option) {
        LinkedHashMap<String, Integer> optionMap = new LinkedHashMap<>();
        String[] strArray = option.replaceAll(":0", "").split(";");
        for (String s : strArray) {
            String[] str = s.split(":");
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
        cellStyle.setFontSize(8);
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
