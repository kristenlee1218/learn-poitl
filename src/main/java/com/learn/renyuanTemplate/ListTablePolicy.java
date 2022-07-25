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

import java.util.*;

/**
 * @author ：Kristen
 * @date ：2022/7/12
 * @description :
 */
public class ListTablePolicy extends AbstractRenderPolicy<Object> {

    //第一种测试情况（信息中心_2020年第1批综合测评统计报表(1655863588387)）(14)
    public static String[] group = new String[]{"对党忠诚", "勇于创新", "治企有方", "兴企有为", "清正廉洁"};
    public static String[][] items = new String[][]{{"政治品质", "政治本领"}, {"创新精神", "创新成果"}, {"经营管理能力", "抓党建强党建能力"}, {"担当作为", "履职绩效"}, {"一岗双责", "廉洁从业"}};
    public static String[][] data = new String[][]{{"刘备", "董事长", "80.00", "1", "78.96", "75.69", "74.98", "73.69", "80.56", "83.46", "89.22", "74.36", "78.25", "85.99", "82.13", "78.93", "79.41", "71.29", "75.48"}, {"关羽", "总经理", "79.00", "2", "78.86", "75.77", "74.39", "73.28", "80.46", "83.46", "89.82", "74.99", "78.26", "85.99", "82.33", "78.93", "84.41", "91.29", "73.48"}};
    public static String[] item = new String[]{""};

    // 第二种测试情况（河北建投二级单位_2022年第1批综合考核评价统计报表(1653979325662)）
//    public static String[] group = new String[]{};
//    public static String[][] items = new String[][]{{"政治品质", "政治本领"}, {"创新精神", "创新成果"}, {"经营管理能力", "抓党建强党建能力"}, {"担当作为", "履职绩效"}, {"一岗双责", "廉洁从业"}};
//    public static String[] item = new String[]{"政治品质", "政治本领", "创新精神", "创新成果", "经营管理能力", "抓党建强党建能力", "担当作为", "履职绩效", "一岗双责", "廉洁从业"};
//    public static String[][] data = new String[][]{{"刘备", "董事长", "80.00", "1", "83.46", "89.22", "74.36", "78.25", "85.99", "82.13", "78.93", "79.41", "71.29", "75.48"}, {"关羽", "总经理", "79.00", "2", "83.46", "89.82", "74.99", "78.26", "85.99", "82.33", "78.93", "84.41", "91.29", "73.48"}};


    // 计算行和列
    int col;
    int row;

    @Override
    public void afterRender(RenderContext<Object> renderContext) {
        // 清空标签
        clearPlaceholder(renderContext, true);
    }

    @Override
    public void doRender(RenderContext<Object> renderContext) {
        XWPFRun run = renderContext.getRun();
        // 当前位置的容器
        BodyContainer bodyContainer = BodyContainerFactory.getBodyContainer(run);

        // 计算行列
        if (group.length > 0) {
            col = this.countCol(items) + 5;
        } else {
            col = item.length + 5;
        }
        row = data.length + 3;

        // 当前位置插入表格
        XWPFTable table = bodyContainer.insertNewTable(run, row, col);
        this.setTableStyle(table);
        this.setTableTitle(table);
        this.setTableHeader(table);
        this.setTableCellData(table);
        //this.setTableCellTag(table);
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

    // 设置第二、三行标题的样式
    public void setTableHeader(XWPFTable table) {
        if (group.length > 0) {
            // 构建第二行的数组、并垂直合并与水平合并
            String[] strHeader1 = new String[5 + group.length];
            strHeader1[0] = "序号";
            strHeader1[1] = "姓名";
            strHeader1[2] = "职务";
            strHeader1[3] = "汇总得分";
            strHeader1[4] = "排名";

            int start = 5;
            int end;
            // 水平合并 group 的名字
            for (String[] value : items) {
                end = (value.length + 1);
                TableTools.mergeCellsHorizonal(table, 1, start, start + end - 1);
                start++;
            }

            for (int i = 0; i < group.length; i++) {
                strHeader1[i + 5] = group[i];
                TableTools.mergeCellsVertically(table, i, 1, 2);
            }

            // 构建第三行的数组
            String[] strHeader2 = new String[col];
            int index = 5;
            for (String[] str : items) {
                strHeader2[index] = "小计";
                for (String s : str) {
                    strHeader2[++index] = s;
                }
                index++;
            }

            // 构建第二、三行
            Style style = this.getCellStyle();
            RowRenderData header1 = this.build(style, strHeader1);
            RowRenderData header2 = this.build(style, strHeader2);
            TableStyle tableStyle = this.getTableStyle();
            header1.setRowStyle(tableStyle);
            header2.setRowStyle(tableStyle);
            MiniTableRenderPolicy.Helper.renderRow(table, 1, header1);
            MiniTableRenderPolicy.Helper.renderRow(table, 2, header2);
        } else {
            // 构建第二行的数组、并垂直合并与水平合并
            String[] strHeader1 = new String[col];
            strHeader1[0] = "序号";
            strHeader1[1] = "姓名";
            strHeader1[2] = "职务";
            strHeader1[3] = "汇总得分";
            strHeader1[4] = "排名";
            int index = 5;
            for (String s : item) {
                strHeader1[index++] = s;
            }
            for (int i = 0; i < col; i++) {
                TableTools.mergeCellsVertically(table, i, 1, 2);
            }
            Style style = this.getCellStyle();
            RowRenderData header1 = this.build(style, strHeader1);
            TableStyle tableStyle = this.getTableStyle();
            header1.setRowStyle(tableStyle);
            MiniTableRenderPolicy.Helper.renderRow(table, 1, header1);
        }
    }

    // 设置行数据
    public void setTableCellData(XWPFTable table) {
        for (int i = 0; i < data.length; i++) {
            String[] str = new String[col];
            str[0] = String.valueOf(i + 1);
            System.arraycopy(data[i], 0, str, 1, data[i].length);
            Style style = this.getDataCellStyle();
            RowRenderData dataRow = this.build(style, str);
            TableStyle tableStyle = this.getTableStyle();
            dataRow.setRowStyle(tableStyle);
            MiniTableRenderPolicy.Helper.renderRow(table, i + 3, dataRow);
        }
    }

//    // 设置数据行的标签
//    public void setTableCellTag(XWPFTable table) {
//        String[] str = new String[col];
//        str[0] = "{{sequence}}";
//        str[1] = "{{leadername}}";
//        str[2] = "{{post}}";
//        str[3] = "{{avg}}";
//        str[4] = "{{sort@avg}}";
//
//        int index = 5;
//        for (int i = 0; i < item.length; i++) {
//            str[index] = "{{avg#group0" + (i + 1) + "}}";
//            for (int j = 0; j < item[i].length; j++) {
//                if (index - 5 - i < 9) {
//                    str[++index] = "{{avg#leader0" + (index - 5 - i) + "}}";
//                } else {
//                    str[++index] = "{{avg#leader" + (index - 5 - i) + "}}";
//                }
//            }
//            index++;
//        }
//        Style style = this.getDataCellStyle();
//        RowRenderData row = this.build(style, str);
//        TableStyle tableStyle = this.getTableStyle();
//        row.setRowStyle(tableStyle);
//        for (int i = 0; i < list.size(); i++) {
//            MiniTableRenderPolicy.Helper.renderRow(table, i + 3, row);
//        }
//    }

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
