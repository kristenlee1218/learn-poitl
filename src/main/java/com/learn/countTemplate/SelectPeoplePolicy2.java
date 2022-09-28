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
import java.util.List;

/**
 * @author ：Kristen
 * @date ：2022/9/26
 * @description : 选人用人 2
 */
public class SelectPeoplePolicy2 extends AbstractRenderPolicy<Object> {

    public static String[] voteType = new String[]{"A1/A2", "A3", "B", "C"};
    public static String[] question = new String[]{"3、您认为本单位选人用人工作存在的主要问题是什么？（可多选）"};
    public static String[] item = new String[]{"1、落实中央关于领导班子和干部队伍建设工作要求有差距",
            "2、选人用人把关不严、质量不高", "3、坚持事业为上不够，不能做到以事择人、人岗相适",
            "4、激励担当作为用人导向不鲜明，论资排辈情况严重", "5、选人用人“个人说了算”",
            "6、任人唯亲、拉帮结派", "7、跑官要官、买官卖官、说情打招呼", "8、执行干部选拔任用政策规定不严格",
            "9、干部队伍建设统筹谋划不够，结构不合理", "10、干部队伍能力素质不适应工作要求"};

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
        col = (voteType.length + 1) * 2 + 1;
        row = item.length + 4;

        // 当前位置插入表格
        XWPFTable table = bodyContainer.insertNewTable(run, row, col);
        this.setTableStyle(table);
        this.setTableTitle(table);
        this.setTableQuestion(table);
        this.setTableHeader(table);
        this.setTableItem(table);
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

    // 设置第二行问题的样式
    public void setTableQuestion(XWPFTable table) {
        Style cellStyle = new Style();
        cellStyle.setFontSize(8);
        cellStyle.setColor("000000");
        cellStyle.setFontFamily("黑体");
        TableStyle tableStyle = new TableStyle();
        String questionName = question[0];
        RowRenderData header1 = RowRenderData.build(new TextRenderData(questionName, cellStyle));
        header1.setRowStyle(tableStyle);
        TableTools.mergeCellsHorizonal(table, 1, 0, col - 1);
        MiniTableRenderPolicy.Helper.renderRow(table, 1, header1);
    }

    public void setTableHeader(XWPFTable table) {
        // 第2行值的数组
        String[] strHeader1 = new String[voteType.length + 2];
        strHeader1[0] = "结果";
        System.arraycopy(voteType, 0, strHeader1, 1, voteType.length);
        strHeader1[strHeader1.length - 1] = "合计";

        // 第2行垂直合并单元格
        TableTools.mergeCellsVertically(table, 0, 2, 3);
        // 第2行水平合并单元格
        int startHeader1 = 1;
        for (int i = 0; i < voteType.length + 1; i++) {
            TableTools.mergeCellsHorizonal(table, 2, startHeader1, startHeader1 + 1);
            startHeader1++;
        }

        // 第3行值的数组 票种部分
        String[] strHeader2 = new String[col];
        for (int i = 0; i < (voteType.length + 1) * 2; i++) {
            if (i % 2 == 1) {
                strHeader2[i + 1] = "票数";
            } else {
                strHeader2[i + 1] = "百分百";
            }
        }


        // 构建第2-3行
        Style style = this.getCellStyle();
        RowRenderData header1 = this.build(strHeader1, style);
        RowRenderData header2 = this.build(strHeader2, style);
        TableStyle tableStyle = this.getTableStyle();
        header1.setRowStyle(tableStyle);
        header2.setRowStyle(tableStyle);
        MiniTableRenderPolicy.Helper.renderRow(table, 2, header1);
        MiniTableRenderPolicy.Helper.renderRow(table, 3, header2);
    }

    // 设置第1列
    public void setTableItem(XWPFTable table) {
        Style cellDataStyle = this.getDataCellStyle();
        TableStyle tableStyle = this.getTableStyle();
        for (int i = 0; i < item.length; i++) {
            RowRenderData itemData = RowRenderData.build(new TextRenderData(item[i], cellDataStyle));
            itemData.setRowStyle(tableStyle);
            MiniTableRenderPolicy.Helper.renderRow(table, i + 4, itemData);
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