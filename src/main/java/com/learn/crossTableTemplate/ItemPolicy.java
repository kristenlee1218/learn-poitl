package com.learn.crossTableTemplate;

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
 * @date ：2022/7/4
 * @description :
 */
public class ItemPolicy extends AbstractRenderPolicy<Object> {

    // 第四种测试情况（河北建投二级单位_2022年第1批综合考核评价统计报表(1653979325662)）
    public static String[] item = new String[]{"政治方向", "社会责任", "企业党建", "科学管理", "发扬民主", "整体合力", "诚信合力", "联系群众", "廉洁自律", "工作部署"};
    public static String[] voteType = new String[]{"A票", "B票"};
    public static String[] value = new String[]{"1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31", "32", "33"};
    public static int year = 2022;
    public static String depart = "信息中心";

    // 计算行
    int row = item.length + 4;
    // 计算列
    int col = voteType.length + 2;

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
        this.setTableStyle(table);
        this.setTableTitle(table);
        this.setTableHeader(table);
        this.setItem(table);
        this.setLastRow(table);
        this.setTableData(table);
    }

    // 整个 table 的样式在此设置
    public void setTableStyle(XWPFTable table) {
        // 设置 A4 幅面的平铺类型和列数
        if (col <= 15) {
            TableTools.widthTable(table, MiniTableRenderData.WIDTH_A4_FULL, col);
        } else {
            TableTools.widthTable(table, MiniTableRenderData.WIDTH_A4_NARROW_FULL, col);
        }
        TableTools.borderTable(table, 10);
        for (XWPFTableRow tableRow : table.getRows()) {
            for (int i = 0; i < tableRow.getTableCells().size(); i++) {
                tableRow.getCell(i).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
                tableRow.getCell(i).setWidth("500");
            }
        }
        table.setCellMargins(5, 10, 5, 10);
        table.setTableAlignment(TableRowAlign.CENTER);
    }

    // 设置第一行标题的样式
    public void setTableTitle(XWPFTable table) {
        Style cellStyle = new Style();
        cellStyle.setFontSize(12);
        cellStyle.setFontFamily("黑体");
        cellStyle.setColor("000000");
        TableStyle tableStyle = new TableStyle();
        tableStyle.setAlign(STJc.CENTER);
        tableStyle.setBackgroundColor("DCDCDC");
        String title = "{{title}}";
        RowRenderData header0 = RowRenderData.build(new TextRenderData(title, cellStyle));
        header0.setRowStyle(tableStyle);
        TableTools.mergeCellsHorizonal(table, 0, 0, col - 1);
        MiniTableRenderPolicy.Helper.renderRow(table, 0, header0);
    }

    // 设置第二、三行标题的样式
    public void setTableHeader(XWPFTable table) {
        int length = voteType.length + 2;
        String[] strHeader1 = new String[length];
        strHeader1[0] = "指标";
        strHeader1[length - 1] = "全体";
        System.arraycopy(voteType, 0, strHeader1, 1, voteType.length);

        // 构建第二行
        Style cellStyle = this.getCellStyle();
        RowRenderData header1 = this.build(strHeader1, cellStyle);

        // 垂直合并
        for (int i = col - 1; i >= 0; i--) {
            TableTools.mergeCellsVertically(table, i, 1, 2);
        }
        TableStyle tableStyle = this.getTableStyle();
        header1.setRowStyle(tableStyle);
        MiniTableRenderPolicy.Helper.renderRow(table, 1, header1);
    }

    // item 的设置、集中在第1列
    public void setItem(XWPFTable table) {
        Style cellStyle = this.getCellStyle();
        TableStyle tableStyle = this.getTableStyle();
        for (int i = 0; i < item.length; i++) {
            RowRenderData itemData = RowRenderData.build(new TextRenderData(item[i], cellStyle));
            itemData.setRowStyle(tableStyle);
            MiniTableRenderPolicy.Helper.renderRow(table, i + 3, itemData);
        }
    }

    // 设置最后一行
    public void setLastRow(XWPFTable table) {
        // 最后一行的加权总得分
        Style cellStyle = this.getCellStyle();
        RowRenderData total = RowRenderData.build(new TextRenderData("加权汇总得分", cellStyle));
        TableStyle tableStyle = this.getTableStyle();
        total.setRowStyle(tableStyle);
        MiniTableRenderPolicy.Helper.renderRow(table, row - 1, total);
    }

    // 设置数据
    public void setTableData(XWPFTable table) {
        Style cellStyle = this.getCellStyle();
        TableStyle tableStyle = this.getTableStyle();
        int start = 3;
        int index = 0;
        for (int i = start; i < row; i++) {
            String[] str = new String[col];
            for (int j = 1; j < table.getRow(i).getTableCells().size(); j++) {
                str[j] = value[index];
                index++;
            }
            RowRenderData lineValue = this.build(str, cellStyle);
            lineValue.setRowStyle(tableStyle);
            MiniTableRenderPolicy.Helper.renderRow(table, i, lineValue);
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
}
