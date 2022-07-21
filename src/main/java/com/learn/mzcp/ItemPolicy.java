package com.learn.mzcp;

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
import com.first.item.Item;
import com.first.voterule.VoteRule;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
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
    private ArrayList<Item> itemList = new ArrayList<>();
    private ArrayList<VoteRule> voterTypeList = new ArrayList<>();
    private String year;
    private String organCode;

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
        row = itemList.size() + 4;
        col = voterTypeList.size() + 2;

        // 当前位置插入表格
        XWPFTable table = bodyContainer.insertNewTable(run, row, col);
        this.setTableStyle(table);
        this.setTableTitle(table);
        this.setTableHeader(table);
        this.setItem(table);
        this.setCellTag(table);
        this.setLastRow(table);
    }

    // 整个 table 的样式在此设置
    public void setTableStyle(XWPFTable table) {
        TableTools.widthTable(table, MiniTableRenderData.WIDTH_A4_FULL, 10);
        TableTools.borderTable(table, 4);
        // 设置表格居中
        TableStyle tableStyle = this.getTableStyle();
        TableTools.styleTable(table, tableStyle);
    }

    // 设置第一行标题的样式
    public void setTableTitle(XWPFTable table) {
        Style cellStyle = new Style();
        cellStyle.setFontSize(12);
        cellStyle.setColor("000000");
        cellStyle.setFontFamily("黑体");
        TableStyle style = new TableStyle();
        style.setAlign(STJc.CENTER);
        style.setBackgroundColor("DCDCDC");
        String title = organCode + "领导班子" + year + "年度综合测评汇总表";
        RowRenderData header0 = RowRenderData.build(new TextRenderData(title));
        header0.setRowStyle(style);
        TableTools.mergeCellsHorizonal(table, 0, 0, col - 1);
        MiniTableRenderPolicy.Helper.renderRow(table, 0, header0);
    }

    // 设置第二、三行标题的样式
    public void setTableHeader(XWPFTable table) {
        int length = this.getVoterTypeList().size() + 2;
        String[] strHeader1 = new String[length];
        strHeader1[0] = "指标";
        strHeader1[length - 1] = "全体";
        for (int i = 0; i < voterTypeList.size(); i++) {
            strHeader1[i + 1] = voterTypeList.get(i).getFilterName() + "（" + voterTypeList.get(i)
                    .getRuleName() + "）";
        }
        // 构建第二行
        Style cellStyle = this.getCellStyle();
        RowRenderData header1 = this.build(cellStyle, strHeader1);

        // 垂直合并
        for (int i = col - 1; i >= 0; i--) {
            TableTools.mergeCellsVertically(table, i, 1, 2);
        }
        MiniTableRenderPolicy.Helper.renderRow(table, 1, header1);
    }

    // item 的设置、集中在第1列
    public void setItem(XWPFTable table) {
        Style cellStyle = this.getCellStyle();
        TableStyle tableStyle = this.getTableStyle();
        for (int i = 0; i < itemList.size(); i++) {
            RowRenderData itemData = RowRenderData.build(new TextRenderData(itemList.get(i).getItemName(), cellStyle));
            itemData.setRowStyle(tableStyle);
            MiniTableRenderPolicy.Helper.renderRow(table, i + 3, itemData);
        }
    }

    // 设置 table 的 cell 中的 tag
    public void setCellTag(XWPFTable table) {
        TableStyle tableStyle = this.getTableStyle();
        Style cellStyle = this.getCellStyle();
        String[] str = new String[col];
        for (int i = 0; i < itemList.size(); i++) {
            Item item = itemList.get(i);
            for (int j = 1; j < str.length; j++) {
                str[j] = "{{avg#" + item.getItemID() + "@" + voterTypeList.get(i - 1) + "}}";
                if (j == str.length - 1) {
                    str[j] = "{{avg#" + item.getItemID() + "}}";
                }
            }
            RowRenderData rowData = this.build(cellStyle, str);
            rowData.setRowStyle(tableStyle);
            MiniTableRenderPolicy.Helper.renderRow(table, i + 3, rowData);
        }
    }

    // 设置最后一行
    public void setLastRow(XWPFTable table) {
        // 最后一行的加权总得分
        String[] str = new String[col];
        str[0] = "加权汇总得分";
        for (int i = 0; i < voterTypeList.size(); i++) {
            str[i + 1] = "{{avg@" + voterTypeList.get(i).getVoterType() + "}}";
        }
        str[str.length - 1] = "{{avg}}";
        Style cellStyle = this.getCellStyle();
        RowRenderData total = this.build(cellStyle, str);
        TableStyle tableStyle = this.getTableStyle();
        total.setRowStyle(tableStyle);
        MiniTableRenderPolicy.Helper.renderRow(table, row - 1, total);
    }

    // 根据 String[] 构建一行的数据，同一行使用一个 Style
    public RowRenderData build(Style style, String... cellStr) {
        List<TextRenderData> data = new ArrayList<>();
        if (null != cellStr) {
            for (String col : cellStr) {
                data.add(new TextRenderData(col, style));
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

    public ArrayList<Item> getItemList() {
        return itemList;
    }

    public void setItemList(ArrayList<Item> itemList) {
        this.itemList = itemList;
    }


    public ArrayList<VoteRule> getVoterTypeList() {
        return voterTypeList;
    }

    public void setVoterTypeList(ArrayList<VoteRule> voterTypeList) {
        this.voterTypeList = voterTypeList;
    }

    public String getYear() {
        return year;
    }

    public void setYear(String year) {
        this.year = year;
    }

    public String getOrganCode() {
        return organCode;
    }

    public void setOrganCode(String organCode) {
        this.organCode = organCode;
    }
}
