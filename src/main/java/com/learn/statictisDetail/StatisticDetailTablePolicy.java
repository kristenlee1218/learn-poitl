package com.learn.statictisDetail;

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
 * @date ：2023/8/21
 * @description : 通用 policy，生成 organ 和 leader 的明细表分解，通过 element 来区分 organ 和 leader，
 * 通过票种来确认有几个 table
 */
public class StatisticDetailTablePolicy extends AbstractRenderPolicy<Object> {

    public static String[] group = new String[]{"政治思想建设", "企业发展质量", "党建工作质量", "作风建设成效"};
    public static String[][] items = new String[][]{{"政治忠诚", "政治担当", "社会责任"}, {"改革创新", "经营效益", "管理效能", "风险管控"}, {"选人用人", "基层党建", "党风廉政"}, {"团结协作", "联系群众"}};
    public static String[] item = new String[]{};
    public static String filterName = "领导班子成员";
    public static String voteType = "A1、A2、A3票";
    public static int score = 10;

//    public static String[] group = new String[]{"政治思想建设", "企业发展质量", "党建工作质量", "作风建设成效"};
//    public static String[][] items = new String[][]{{"政治忠诚", "政治担当", "社会责任"}, {"改革创新", "经营效益", "管理效能", "风险管控"}, {"选人用人", "基层党建", "党风廉政"}, {"团结协作", "联系群众"}};
//    public static String[] item = new String[]{};
//    public static String[] voteType = new String[]{"A1、A2、A3票"};
//    public static String scoreGrade = "10";

    // 标题+表头+加权总数所占的固定行数
    int rowBase = 4;
    // “得票” +评分的固定列数
    int colBase = 3;

    // 计算行
    int row;
    // 计算列
    int col;

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
        row = this.getRow(group);
        col = this.getCol(group);

        // 当前位置插入表格
        XWPFTable table = bodyContainer.insertNewTable(run, row, col);

        this.setTableStyle(table);
        this.setTableTitle(table);
        this.setTableHeader(table);
//        this.setItem(table);
//        this.setLastRow(table);

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

    public void setTableHeader(XWPFTable table) {
        if (group.length > 0) {
            // 构建第二行
            String[] strHeader1 = new String[2];
            strHeader1[0] = "评价内容";
            strHeader1[1] = filterName + "（" + voteType + "）";
            for (int i = 0; i < 3; i++) {
                TableTools.mergeCellsVertically(table, i, 1, 2);
            }
            TableTools.mergeCellsHorizonal(table, 1, 3, col - 1);
            TableTools.mergeCellsHorizonal(table, 1, 0, 2);
            TableTools.mergeCellsHorizonal(table, 2, col - 2, col - 1);
            TableTools.mergeCellsHorizonal(table, 2, 0, 2);

            String[] strHeader2 = new String[score + 2];
            for (int i = score; i > 0; i--) {
                strHeader2[strHeader2.length - i - 1] = String.valueOf(i);
            }
            strHeader2[strHeader2.length-1] = "评分";


            Style cellStyle = this.getCellStyle();
            RowRenderData header1 = this.build(strHeader1, cellStyle);
            RowRenderData header2 = this.build(strHeader2, cellStyle);
            TableStyle tableStyle = this.getTableStyle();
            header1.setRowStyle(tableStyle);
            header2.setRowStyle(tableStyle);
            MiniTableRenderPolicy.Helper.renderRow(table, 1, header1);
            MiniTableRenderPolicy.Helper.renderRow(table, 2, header2);

        }
    }

    // 计算行（区分有没有一级指标）
    public int getRow(String[] group) {
        // 有一级指标和二级指标
        if (group.length > 0) {
            row = this.countRow(items) + rowBase;
        } else {
            row = item.length + rowBase;
        }
        return row;
    }

    // 计算所有分组的项的个数
    public int countRow(String[][] str) {
        int count = 0;
        for (String[] s : str) {
            for (int j = 0; j < s.length; j++) {
                count++;
            }
        }
        return count;
    }

    // 计算列（区分有没有一级指标）
    public int getCol(String[] group) {
        // 有一级指标和二级指标
        if (group.length > 0) {
            col = score + colBase + 2;
        } else {
            col = score + colBase + 1;
        }
        return col;
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

    // 设置 cell 格样式
    public Style getDataCellStyle() {
        Style cellStyle = new Style();
        cellStyle.setFontFamily("宋体");
        cellStyle.setFontSize(8);
        cellStyle.setColor("000000");
        return cellStyle;
    }
}
