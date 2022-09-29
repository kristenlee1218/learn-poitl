package com.learn.countTemplate;

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
 * @date ：2022/9/26
 * @description : 新提拔
 */
public class NewLeaderPolicy extends AbstractRenderPolicy<Object> {

    public static String[] voteType = new String[]{"A1", "A2", "A3", "B", "C"};
    public static String option = "4:不了解:0;6:不认同:0;8:基本认同:0;10:认同:0";
    public static String[] question = new String[]{"对提拔任用该干部的看法"};
    public static String[] properties = new String[]{"序号", "姓名", "出生年月", "原任职务", "现任职务"};
    public static String[] data = new String[]{};

    //public static String[] voteType = new String[]{"A1/A2", "A3", "B", "C"};
    //public static String[] properties = new String[]{"序号", "姓名", "出生年月", "原任职务", "现任职务", "任职时间"};

    // 计算行和列
    int col;
    int row;
    LinkedHashMap<String, Integer> optionMap;

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
        row = 5 + data.length;
        col = properties.length + voteType.length * optionMap.size() + (optionMap.size() + 1) * 2;
        // 当前位置插入表格
        XWPFTable table = bodyContainer.insertNewTable(run, row, col);
        this.setTableStyle(table);
        this.setTableTitle(table);
        this.setTableHeader(table);
        //this.setTableTag(table);
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
        // 第1行值的数组
        String[] strHeader1 = new String[2];
        strHeader1[0] = "被评议对象的基本情况";
        strHeader1[1] = question[0];

        // 第1行水平合并单元格 对提拔任用该干部的看法部分
        TableTools.mergeCellsHorizonal(table, 1, properties.length, col - 1);
        // 第1行水平合并单元格 被评议对象的基本情况部分
        TableTools.mergeCellsHorizonal(table, 1, 0, properties.length - 1);
        TableTools.mergeCellsHorizonal(table, 2, 0, properties.length - 1);

        // 第1行垂直合并单元格
        TableTools.mergeCellsVertically(table, 0, 1, 2);

        // 第2行值的数组
        String[] strHeader2 = new String[voteType.length + 2];
        System.arraycopy(voteType, 0, strHeader2, 1, voteType.length);
        strHeader2[strHeader2.length - 1] = "合计";

        // 第2行水平合并单元格 票种部分
        int startHeader2 = 1;
        int endHeader2;
        for (int i = 0; i < strHeader2.length - 2; i++) {
            endHeader2 = startHeader2 + optionMap.size() - 1;
            TableTools.mergeCellsHorizonal(table, 2, startHeader2, endHeader2);
            startHeader2++;
        }

        // 第2行水平合并单元格 合计部分
        TableTools.mergeCellsHorizonal(table, 2, startHeader2, startHeader2 + 2 * (optionMap.size() + 1) - 1);

        // 第3行值的数组
        String[] strHeader3 = new String[properties.length + (optionMap.size()) * (voteType.length + 1) + 1];
        System.arraycopy(properties, 0, strHeader3, 0, properties.length);
        int header3Index = properties.length;
        for (int i = 0; i < voteType.length + 1; i++) {
            for (int j = 0; j < optionMap.keySet().toArray().length; j++) {
                strHeader3[header3Index++] = optionMap.keySet().toArray()[j].toString();
            }
        }
        strHeader3[strHeader3.length - 1] = "认同度";

        // 第3行水平合并单元格 合计部分
        int startHeader3 = voteType.length * optionMap.size() + properties.length;
        for (int i = 0; i < optionMap.size(); i++) {
            TableTools.mergeCellsHorizonal(table, 3, startHeader3, startHeader3 + 1);
            startHeader3++;
        }

        // 第3行水平合并单元格 认同度部分
        TableTools.mergeCellsHorizonal(table, 3, startHeader3, startHeader3 + 1);

        // 第3行垂直合并单元格 属性部分
        for (int i = 0; i < properties.length; i++) {
            TableTools.mergeCellsVertically(table, i, 3, 4);
        }

        // 第4行数组的值
        String[] strHeader4 = new String[col];
        for (int i = 0; i < voteType.length * optionMap.size(); i++) {
            strHeader4[i + properties.length] = "得票";
        }

        // 第4行值的数组 合计部分
        for (int i = voteType.length * optionMap.size() + properties.length; i < strHeader4.length; i++) {
            if ((voteType.length * optionMap.size() + properties.length) % 2 == 0) {
                if (i % 2 == 0) {
                    strHeader4[i] = "得票";
                } else {
                    strHeader4[i] = "比例";
                }
            } else {
                if (i % 2 == 0) {
                    strHeader4[i] = "比例";
                } else {
                    strHeader4[i] = "得票";
                }
            }
        }
        strHeader4[strHeader4.length - 1] = "排名";

        // 构建第1-4行
        Style style = this.getCellStyle();
        RowRenderData header1 = this.build(strHeader1, style);
        RowRenderData header2 = this.build(strHeader2, style);
        RowRenderData header3 = this.build(strHeader3, style);
        RowRenderData header4 = this.build(strHeader4, style);
        TableStyle tableStyle = this.getTableStyle();
        header1.setRowStyle(tableStyle);
        header2.setRowStyle(tableStyle);
        header3.setRowStyle(tableStyle);
        header4.setRowStyle(tableStyle);
        MiniTableRenderPolicy.Helper.renderRow(table, 1, header1);
        MiniTableRenderPolicy.Helper.renderRow(table, 2, header2);
        MiniTableRenderPolicy.Helper.renderRow(table, 3, header3);
        MiniTableRenderPolicy.Helper.renderRow(table, 4, header4);
    }

    // 将题目的选项 如：“4:不了解:0;6:不认同:0;8:基本认同:0;10:认同:0” 存入 map，key 为显示的值，value 为分值
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
