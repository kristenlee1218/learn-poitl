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
    public static String[] item = new String[]{"政治忠诚", "政治担当", "社会责任","改革创新", "经营效益", "管理效能", "风险管控","选人用人", "基层党建", "党风廉政","团结协作", "联系群众"};
    public static String filterName = "领导班子成员";
    public static String voteRuleName = "A1、A2、A3票";
    public static String votertype = "A1;A2;A3";
    public static int score = 10;

//    public static String[] group = new String[]{"对党忠诚", "勇于创新", "治企有方", "兴企有为", "清正廉洁"};
//    public static String[][] items = new String[][]{{"政治品质", "政治本领"}, {"创新精神", "创新成果"}, {"经营管理能力", "抓党建强党建能力"}, {"担当作为", "履职绩效"}, {"一岗双责", "廉洁从业"}};
//    public static String[] item = new String[]{"政治品质", "政治本领", "创新精神", "创新成果", "经营管理能力", "抓党建强党建能力", "担当作为", "履职绩效", "一岗双责", "廉洁从业"};
//    public static String[] voteType = new String[]{"A1、A2、A3票"};
//    public static String filterName = "领导班子成员";
//    public static String voteRuleName = "A1、A2、A3票";
//    public static String votertype = "A1;A2;A3";
//    public static int score = 10;

//    public static String[] group = new String[]{};
//    public static String[][] items = new String[][]{{"政治品质", "政治本领"}, {"创新精神", "创新成果"}, {"经营管理能力", "抓党建强党建能力"}, {"担当作为", "履职绩效"}, {"一岗双责", "廉洁从业"}};
//    public static String[] item = new String[]{"政治品质", "政治本领", "创新精神", "创新成果", "经营管理能力", "抓党建强党建能力", "担当作为", "履职绩效", "一岗双责", "廉洁从业"};
//    public static String[] voteType = new String[]{"A1、A2、A3票"};
//    public static String filterName = "领导班子成员";
//    public static String voteRuleName = "A1、A2、A3票";
//    public static String votertype = "A1;A2;A3";
//    public static int score = 10;


    // 标题+表头+加权总数所占的固定行数
    int rowBase = 3;
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
        if (group.length > 0) {
            this.setGroup(table);
        }
        this.setItem(table);
        this.setText(table);
        this.setLastRow(table);
        this.setCellTag(table);
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
        tableStyle.setBackgroundColor("E7E6E6");
        String title = "{{title}}";
        RowRenderData header0 = RowRenderData.build(new TextRenderData(title, cellStyle));
        header0.setRowStyle(tableStyle);
        TableTools.mergeCellsHorizonal(table, 0, 0, col - 1);
        MiniTableRenderPolicy.Helper.renderRow(table, 0, header0);
    }

    // 设置 header
    public void setTableHeader(XWPFTable table) {
        if (group.length > 0) {
            for (int i = 0; i < 3; i++) {
                TableTools.mergeCellsVertically(table, i, 1, 2);
            }
            TableTools.mergeCellsHorizonal(table, 1, 3, col - 1);
            TableTools.mergeCellsHorizonal(table, 1, 0, 2);
            TableTools.mergeCellsHorizonal(table, 2, col - 2, col - 1);
            TableTools.mergeCellsHorizonal(table, 2, 0, 2);
        } else {
            for (int i = 0; i < 2; i++) {
                TableTools.mergeCellsVertically(table, i, 1, 2);
            }
            TableTools.mergeCellsHorizonal(table, 1, 2, col - 1);
            TableTools.mergeCellsHorizonal(table, 1, 0, 1);
            TableTools.mergeCellsHorizonal(table, 2, 0, 1);
        }
        // 构建第二行
        String[] strHeader1 = new String[2];
        strHeader1[0] = "评价内容";
        strHeader1[1] = filterName + "（" + voteRuleName + "）";
        String[] strHeader2 = new String[score + 2];
        for (int i = score; i > 0; i--) {
            strHeader2[strHeader2.length - i - 1] = String.valueOf(i);
        }
        strHeader2[strHeader2.length - 1] = "评分";
        Style cellStyle = this.getCellStyle();
        RowRenderData header1 = this.build(strHeader1, cellStyle);
        RowRenderData header2 = this.build(strHeader2, cellStyle);
        TableStyle tableStyle = this.getTableStyle();
        header1.setRowStyle(tableStyle);
        header2.setRowStyle(tableStyle);
        MiniTableRenderPolicy.Helper.renderRow(table, 1, header1);
        MiniTableRenderPolicy.Helper.renderRow(table, 2, header2);
    }

    // group 的设置、集中在第1列
    public void setGroup(XWPFTable table) {
        int start = rowBase;
        int end;
        Style cellStyle = this.getCellStyle();
        TableStyle tableStyle = this.getTableStyle();
        String[] str = new String[col];
        for (int i = 0; i < items.length; i++) {
            end = start + items[i].length;
            TableTools.mergeCellsVertically(table, 0, start, end - 1);
            TableTools.mergeCellsVertically(table, col - 1, start, end - 1);
            RowRenderData groupData = RowRenderData.build(new TextRenderData(group[i], cellStyle));
            str[col - 1] = "count#group0" + (i + 1) + "@" + votertype;
            RowRenderData groupTag = this.build(str, cellStyle);
            groupData.setRowStyle(tableStyle);
            MiniTableRenderPolicy.Helper.renderRow(table, start, groupData);
            MiniTableRenderPolicy.Helper.renderRow(table, start, groupTag);
            start = end;
        }
    }

    // item 的设置、集中在第2列
    public void setItem(XWPFTable table) {
        // 有一级指标并含有分组
        if (group.length > 0) {
            // item 的设置、集中在第 2 列
            // 构建 item 列
            int index = rowBase;
            Style cellStyle = this.getCellStyle();
            TableStyle tableStyle = this.getTableStyle();
            for (String[] str : items) {
                for (String s : str) {
                    RowRenderData itemData = RowRenderData.build(new TextRenderData("", cellStyle), new TextRenderData(s, cellStyle));
                    itemData.setRowStyle(tableStyle);
                    MiniTableRenderPolicy.Helper.renderRow(table, index++, itemData);
                }
            }
        } else {
            TableStyle tableStyle = this.getTableStyle();
            Style cellStyle = this.getCellStyle();
            for (int i = 0; i < item.length; i++) {
                RowRenderData itemData = RowRenderData.build(new TextRenderData(item[i], cellStyle));
                itemData.setRowStyle(tableStyle);
                MiniTableRenderPolicy.Helper.renderRow(table, i + rowBase, itemData);
            }
        }
    }

    // 设置 “得票”
    public void setText(XWPFTable table) {
        int start = rowBase;
        int end = row - 1;
        Style cellStyle = this.getCellStyle();
        TableStyle tableStyle = this.getTableStyle();
        RowRenderData textData;
        if (group.length > 0) {
            textData = RowRenderData.build(new TextRenderData("", cellStyle), new TextRenderData("", cellStyle), new TextRenderData("得票", cellStyle));
        } else {
            textData = RowRenderData.build(new TextRenderData("", cellStyle), new TextRenderData("得票", cellStyle));
        }
        for (int i = start; i < end; i++) {
            textData.setRowStyle(tableStyle);
            MiniTableRenderPolicy.Helper.renderRow(table, i, textData);
        }
    }

    // 设置最后一行
    public void setLastRow(XWPFTable table) {
        // 最后一行的加权总得分
        TableTools.mergeCellsHorizonal(table, row - 1, col - 2, col - 1);
        if (group.length > 0) {
            TableTools.mergeCellsHorizonal(table, row - 1, colBase, col - 3);
            TableTools.mergeCellsHorizonal(table, row - 1, 0, colBase - 1);
        } else {
            TableTools.mergeCellsHorizonal(table, row - 1, 2, col - 3);
            TableTools.mergeCellsHorizonal(table, row - 1, 0, 1);
        }
        Style cellStyle = this.getCellStyle();
        RowRenderData total = RowRenderData.build(new TextRenderData("加权汇总得分", cellStyle), new TextRenderData("——————————", cellStyle), new TextRenderData("avg@" + votertype, cellStyle));
        TableStyle tableStyle = this.getTableStyle();
        total.setRowStyle(tableStyle);
        MiniTableRenderPolicy.Helper.renderRow(table, row - 1, total);
    }

    // 设置 tag
    public void setCellTag(XWPFTable table) {
        Style cellStyle = this.getCellStyle();
        TableStyle tableStyle = this.getTableStyle();
        String[] str = new String[col];
        if (group.length > 0) {
            for (int i = 1; i <= item.length; i++) {
                for (int j = 1; j <= score; j++) {
                    if (i >= 10) {
                        str[j + colBase - 1] = "count#organ" + i + "=" + (10 - j + 1) + "#@" + votertype;
                    } else {
                        str[j + colBase - 1] = "count#organ0" + i + "=" + (10 - j + 1) + "#@" + votertype;
                    }
                }
                str[score + colBase] = "count#organ" + i + "#@" + votertype;
                RowRenderData tag = this.build(str, cellStyle);
                tag.setRowStyle(tableStyle);
                MiniTableRenderPolicy.Helper.renderRow(table, i + rowBase - 1, tag);
            }
        } else {
            for (int i = 1; i <= item.length; i++) {
                for (int j = 1; j <= score; j++) {
                    if (i >= 10) {
                        str[j + 1] = "count#organ" + i + "=" + (10 - j + 1) + "#@" + votertype;
                    } else {
                        str[j + 1] = "count#organ0" + i + "=" + (10 - j + 1) + "#@" + votertype;
                    }
                }
                str[score + 2] = "count#organ" + i + "#@" + votertype;
                RowRenderData tag = this.build(str, cellStyle);
                tag.setRowStyle(tableStyle);
                MiniTableRenderPolicy.Helper.renderRow(table, i + rowBase - 1, tag);
            }
        }
    }

    // 计算行（区分有没有一级指标）
    public int getRow(String[] group) {
        // 有一级指标和二级指标
        row = this.countRow(items) + rowBase + 1;
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
            col = score + colBase;
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
