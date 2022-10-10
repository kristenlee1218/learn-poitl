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

import java.util.*;

/**
 * @author ：Kristen
 * @date ：2022/9/26
 * @description : 选人用人 1
 * 当两个题目的评分值均为 4:不了解:0;6:不好:0;8:一般:0;10:好:0、则将两个题合成一个表格
 * <p>
 * 如果两个题目的评分值不一致、则模版需要写两个问题的标记、两个问题分别调用本 policy，则会生成两个表格
 * <p>
 * ¤1¤2¤3¤5¤6¤7¤
 * <p>
 * concat('¤',str_organ23,'¤') like '%¤6¤%'
 */
public class SelectPeoplePolicy1 extends AbstractRenderPolicy<Object> {

    public static String[] voteType = new String[]{"A1/A2", "A3", "B", "C"};
    public static String option = "4:不了解:0;6:不好:0;8:一般:0;10:好:0";
    public static String[] question = new String[]{"1、对本单位选人用人工作的总体评价", "2、对本单位从严管理监督干部情况的评价"};
    public static String[] itemId = new String[]{"organ21", "organ22"};
    public static String[][] data = new String[][]{{"1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25"},
            {"1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25"}};

    //    public sta tic String[] voteType = new String[]{"A1", "A2", "A3", "B", "C"};
//    public static String option = "2:极差:0;4:不了解:0;6:不好:0;8:一般:0;10:好:0";
//    public static String[] question = new String[]{"1、对本单位选人用人工作的总体评价"};
//    public static String[] itemId = new String[]{"organ21"};

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
        row = 4 + question.length;
        col = voteType.length * optionMap.size() + optionMap.size() * 2 + 2;

        // 当前位置插入表格
        XWPFTable table = bodyContainer.insertNewTable(run, row, col);
        this.setTableStyle(table);
        this.setTableTitle(table);
        this.setTableHeader(table);
        this.setTableQuestion(table);
        //this.setTableTag(table);
        this.setTableData(table);
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
        String[] strHeader1 = new String[voteType.length + 3];
        strHeader1[0] = "结果";
        System.arraycopy(voteType, 0, strHeader1, 1, voteType.length);
        strHeader1[voteType.length + 1] = "合计";
        strHeader1[strHeader1.length - 1] = "好与一般的比例";

        // 第1行垂直合并单元格
        TableTools.mergeCellsVertically(table, 0, 1, 3);
        TableTools.mergeCellsVertically(table, col - 1, 1, 3);

        // 第1行水平合并单元格 票种部分
        int startHeader1 = 1;
        int endHeader1;
        for (int i = 1; i < strHeader1.length - 2; i++) {
            endHeader1 = startHeader1 + optionMap.size() - 1;
            TableTools.mergeCellsHorizonal(table, 1, startHeader1, endHeader1);
            startHeader1++;
        }
        // 第1行水平合并单元格 合计部分
        TableTools.mergeCellsHorizonal(table, 1, startHeader1, startHeader1 + 2 * optionMap.size() - 1);

        // 第2行值的数组
        String[] strHeader2 = new String[(voteType.length + 1) * optionMap.size() + 1];
        int header2Index = 1;
        for (int i = 0; i < voteType.length + 1; i++) {
            for (int j = 0; j < optionMap.keySet().toArray().length; j++) {
                strHeader2[header2Index++] = optionMap.keySet().toArray()[j].toString();
            }
        }

        // 第2行水平合并单元格 合计部分
        int startHeader2 = voteType.length * optionMap.size() + 1;
        for (int i = 0; i < optionMap.size(); i++) {
            TableTools.mergeCellsHorizonal(table, 2, startHeader2, startHeader2 + 1);
            startHeader2++;
        }

        // 第3行值的数组 票种部分
        String[] strHeader3 = new String[col - 1];
        for (int i = 0; i < voteType.length * optionMap.size(); i++) {
            strHeader3[i + 1] = "得票";
        }

        // 第3行值的数组 合计部分
        for (int i = voteType.length * optionMap.size() + 1; i < strHeader3.length; i++) {
            if ((voteType.length * optionMap.size()) % 2 == 0) {
                if (i % 2 == 0) {
                    strHeader3[i] = "比例";
                } else {
                    strHeader3[i] = "得票";
                }
            } else {
                if (i % 2 == 0) {
                    strHeader3[i] = "得票";
                } else {
                    strHeader3[i] = "比例";
                }
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

    // 设置第1列
    public void setTableQuestion(XWPFTable table) {
        Style cellStyle = this.getCellStyle();
        TableStyle tableStyle = this.getTableStyle();
        for (int i = 0; i < question.length; i++) {
            RowRenderData questionData = RowRenderData.build(new TextRenderData(question[i], cellStyle));
            questionData.setRowStyle(tableStyle);
            MiniTableRenderPolicy.Helper.renderRow(table, i + 4, questionData);
        }
    }

    public void setTableTag(XWPFTable table) {
        for (int i = 0; i < question.length; i++) {
            // 设置 tag（票种类型部分）
            String[] strTag = new String[col];
            int index = 1;
            for (String s : voteType) {
                for (int k = 0; k < optionMap.size(); k++) {
                    strTag[index] = "count_" + itemId[i] + "_" + optionMap.values().toArray()[k].toString() + "__" + s.replaceAll("/", "_");
                    index++;
                }
            }
            // 设置 tag（合计部分）
            for (int j = 0; j < optionMap.values().toArray().length; j++) {
                strTag[index++] = "count_" + itemId[i] + "_" + optionMap.values().toArray()[j].toString() + "_";
                strTag[index++] = "rate_" + itemId[i] + "_" + optionMap.values().toArray()[j].toString() + "_";
            }

            // 设置 tag（最后一行部分）
            strTag[col - 1] = "rate_" + itemId[i] + "_7_";

            // 构建
            Style style = this.getCellStyle();
            RowRenderData tag = this.build(strTag, style);
            TableStyle tableStyle = this.getTableStyle();
            tag.setRowStyle(tableStyle);
            MiniTableRenderPolicy.Helper.renderRow(table, i + 4, tag);
        }
    }

    public void setTableData(XWPFTable table) {
        for (int i = 0; i < data.length; i++) {
            String[] strData = new String[col];
            System.arraycopy(data[i], 0, strData, 1, data[i].length);
            // 构建
            Style style = this.getCellStyle();
            RowRenderData tag = this.build(strData, style);
            TableStyle tableStyle = this.getTableStyle();
            tag.setRowStyle(tableStyle);
            MiniTableRenderPolicy.Helper.renderRow(table, i + 4, tag);
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
