package com.learn.crossTableTemplate;

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
 * @date ：2024/8/26
 * @description :
 */
public class OrganItemLineTablePolicy extends AbstractRenderPolicy<Object> {

    public static String[] item = new String[]{"政治忠诚", "政治担当", "社会责任", "改革创新", "经营效益", "管理效能", "风险管控", "选人用人", "基层党建", "党风廉政"};
    public static String[] itemId = new String[]{"organ01", "organ02", "organ03", "organ04", "organ05", "organ06", "organ07", "organ08", "organ09", "organ10"};
    public static String[] votertype = new String[]{"A1", "A2", "A3", "A4", "B", "C"};

    int col = item.length + 2;
    int row = votertype.length + 1;

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
        XWPFTable table = bodyContainer.insertNewTable(run, row, col);

        // 当前位置插入表格
        this.setTableStyle(table);
        this.setItem(table);
        this.setCellTag(table);
    }

    // 整个 table 的样式在此设置
    public void setTableStyle(XWPFTable table) {
        // 设置 A4 幅面的平铺类型和列数
        TableTools.widthTable(table, 18, col);

        // 设置 border
        TableTools.borderTable(table, 10);
        for (XWPFTableRow tableRow : table.getRows()) {
            for (int i = 0; i < tableRow.getTableCells().size(); i++) {
                tableRow.getCell(i).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
                if (i == 0) {
                    tableRow.getCell(i).setWidth("300");
                } else {
                    tableRow.getCell(i).setWidth("15");
                }
            }
            tableRow.setHeight(200);
        }
        table.setCellMargins(2, 2, 2, 2);
        table.setTableAlignment(TableRowAlign.CENTER);
    }

    // 设置 itemName
    public void setItem(XWPFTable table) {
        String[] strHeader0 = new String[col];
        System.arraycopy(item, 0, strHeader0, 1, item.length);
        strHeader0[strHeader0.length - 1] = "各级均值";
        Style style = this.getCellStyle();
        RowRenderData header1 = this.build(strHeader0, style);
        TableStyle tableStyle = this.getTableStyle();
        header1.setRowStyle(tableStyle);
        MiniTableRenderPolicy.Helper.renderRow(table, 0, header1);
    }

    // 设置标记
    public void setCellTag(XWPFTable table) {
        for (int i = 1; i < row; i++) {
            String[] str = new String[col];
            str[0] = votertype[i - 1] + "票评价";
            for (int j = 0; j < itemId.length; j++) {
                str[j + 1] = "{{avg_" + itemId[j] + "_" + votertype[i - 1] + "}}";
            }
            str[str.length - 1] = "{{avg_" + votertype[i - 1] + "}}";
            Style style = this.getCellStyle();
            RowRenderData header = this.build(str, style);
            TableStyle tableStyle = this.getTableStyle();
            header.setRowStyle(tableStyle);
            MiniTableRenderPolicy.Helper.renderRow(table, i, header);
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
        cellStyle.setFontFamily("仿宋");
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
