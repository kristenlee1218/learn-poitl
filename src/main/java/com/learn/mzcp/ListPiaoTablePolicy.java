///*
// * Copyright (c) 2006- 2022, cuitianxin. All Rights Reserved.
// */
//
//package com.learn.mzcp;
//
//import com.deepoove.poi.data.MiniTableRenderData;
//import com.deepoove.poi.data.RowRenderData;
//import com.deepoove.poi.data.TextRenderData;
//import com.deepoove.poi.data.style.Style;
//import com.deepoove.poi.data.style.TableStyle;
//import com.deepoove.poi.policy.AbstractRenderPolicy;
//import com.deepoove.poi.policy.MiniTableRenderPolicy;
//import com.deepoove.poi.render.RenderContext;
//import com.deepoove.poi.util.TableTools;
//import com.deepoove.poi.xwpf.BodyContainer;
//import com.deepoove.poi.xwpf.BodyContainerFactory;
//import com.first.leaderfieldconfig.LeaderFieldConfig;
//import com.first.leaderfieldconfig.LeaderFieldConfigDAO;
//import com.first.voterule.VoteRule;
//import com.first.voterule.VoteRuleDAO;
//import com.first.voterule.VoteRuleSystem;
//import org.apache.poi.xwpf.usermodel.*;
//import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJc;
//
//import java.util.ArrayList;
//import java.util.List;
//
///**
// * @author ：Kristen
// * @date ：2022/8/16
// * @description :
// */
//public class ListPiaoTablePolicy extends AbstractRenderPolicy<Object> {
//
//    private VoteRuleSystem system;
//    private int dataSize;
//
//    // 所有的 voterule
//    public ArrayList<VoteRule> voterTypeList;
//    // 读取配置 list
//    public ArrayList<LeaderFieldConfig> configList;
//
//
//    // 计算行和列
//    int col;
//    int row;
//
//    @Override
//    public void afterRender(RenderContext<Object> renderContext) {
//        // 清空标签
//        clearPlaceholder(renderContext, true);
//    }
//
//    @Override
//    public void doRender(RenderContext<Object> renderContext) throws Exception {
//        XWPFRun run = renderContext.getRun();
//        // 当前位置的容器
//        BodyContainer bodyContainer = BodyContainerFactory.getBodyContainer(run);
//        // 所有的 voterule
//        voterTypeList = VoteRuleDAO.getVoteRuleListBySystem(system.getSystemID());
//        String backSQL = "where menuIndex = '2' and usedelements is not null";
//        configList = LeaderFieldConfigDAO.getLeaderFieldConfigListBySQL(backSQL);
//
//        col = (voterTypeList.size() + 1) * 2 + configList.size() + 1;
//        row = dataSize + 3;
//
//        // 当前位置插入表格
//        XWPFTable table = bodyContainer.insertNewTable(run, row, col);
//        this.setTableStyle(table);
//        this.setTableTitle(table);
//        this.setTableHeader(table);
//        this.setTableCellTag(table);
//    }
//
//    // 整个 table 的样式在此设置
//    public void setTableStyle(XWPFTable table) {
//        // 设置 A4 幅面的平铺类型和列数
//        TableTools.widthTable(table, MiniTableRenderData.WIDTH_A4_MEDIUM_FULL, col);
//        // 设置 border
//        TableTools.borderTable(table, 10);
//        for (XWPFTableRow tableRow : table.getRows()) {
//            for (int i = 0; i < tableRow.getTableCells().size(); i++) {
//                tableRow.getCell(i).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
//                tableRow.getCell(i).setWidth("1000");
//            }
//        }
//        table.setCellMargins(2, 2, 2, 2);
//        table.setTableAlignment(TableRowAlign.CENTER);
//    }
//
//    // 设置第一行标题的样式
//    public void setTableTitle(XWPFTable table) {
//        Style cellStyle = new Style();
//        cellStyle.setFontSize(12);
//        cellStyle.setColor("000000");
//        cellStyle.setFontFamily("黑体");
//        TableStyle tableStyle = new TableStyle();
//        tableStyle.setAlign(STJc.CENTER);
//        tableStyle.setBackgroundColor("DCDCDC");
//        String title = "{{title}}";
//        RowRenderData header = RowRenderData.build(new TextRenderData(title, cellStyle));
//        header.setRowStyle(tableStyle);
//        TableTools.mergeCellsHorizonal(table, 0, 0, col - 1);
//        MiniTableRenderPolicy.Helper.renderRow(table, 0, header);
//    }
//
//    // 设置第二、三行标题的样式
//    public void setTableHeader(XWPFTable table) {
//        // 构建第二行的数组、并垂直合并与水平合并
//        String[] strHeader1 = new String[voterTypeList.size() + configList.size() + 2];
//        strHeader1[0] = "序号";
//        for (int i = 0; i < configList.size(); i++) {
//            strHeader1[i + 1] = configList.get(i).getFieldName();
//        }
//        strHeader1[strHeader1.length - 1] = "全体";
//
//        int index = configList.size() + 1;
//        int start = index;
//        // 水平合并 voteType 的名字
//        for (int i = 0; i < voterTypeList.size() + 1; i++) {
//            TableTools.mergeCellsHorizonal(table, 1, start, ++start);
//        }
//
//        for (int i = 0; i < voterTypeList.size(); i++) {
//            strHeader1[index + i] = voterTypeList.get(i).getShortName();
//        }
//
//        for (int i = 0; i < index; i++) {
//            TableTools.mergeCellsVertically(table, i, 1, 2);
//        }
//
//        // 构建第三行的数组
//        String[] strHeader2 = new String[col];
//        for (int i = 0; i < (voterTypeList.size() + 1) * 2; i++) {
//            if (i % 2 == 0) {
//                strHeader2[index + i] = "得分";
//            } else {
//                strHeader2[index + i] = "排名";
//            }
//        }
//
//        // 构建第二、三行
//        Style style = this.getCellStyle();
//        RowRenderData header1 = this.build(strHeader1, style);
//        RowRenderData header2 = this.build(strHeader2, style);
//        TableStyle tableStyle = this.getTableStyle();
//        header1.setRowStyle(tableStyle);
//        header2.setRowStyle(tableStyle);
//        MiniTableRenderPolicy.Helper.renderRow(table, 1, header1);
//        MiniTableRenderPolicy.Helper.renderRow(table, 2, header2);
//    }
//
//    // 设置数据行的标签
//    public void setTableCellTag(XWPFTable table) {
//        for (int k = 0; k < dataSize; k++) {
//            String[] str = new String[col];
//            str[0] = "{{sequence_" + k + "}}";
//            for (int i = 0; i < configList.size(); i++) {
//                str[i + 1] = "{{" + configList.get(i).getFieldID() + "_" + k + "}}";
//            }
//
//            int index = configList.size() + 1;
//            for (VoteRule voteRule : voterTypeList) {
//                str[index] = "{{avg_" + voteRule.getShortName() + "_" + k + "}}";
//                str[++index] = "{{sort_avg_" + voteRule.getShortName() + "_" + k + "}}";
//                index++;
//            }
//
//            str[str.length - 2] = "{{avg_" + k + "}}";
//            str[str.length - 1] = "{{sort_avg_" + k + "}}";
//            Style style = this.getDataCellStyle();
//            RowRenderData row = this.build(str, style);
//            TableStyle tableStyle = this.getTableStyle();
//            row.setRowStyle(tableStyle);
//            MiniTableRenderPolicy.Helper.renderRow(table, k + 3, row);
//        }
//    }
//
//    // 根据 String[] 构建一行的数据，同一行使用一个 Style
//    public RowRenderData build(String[] cellStr, Style style) {
//        List<TextRenderData> data = new ArrayList<>();
//        if (null != cellStr) {
//            for (String s : cellStr) {
//                data.add(new TextRenderData(s, style));
//            }
//        }
//        return new RowRenderData(data, null);
//    }
//
//    // 设置 cell 格样式
//    public Style getCellStyle() {
//        Style cellStyle = new Style();
//        cellStyle.setFontFamily("宋体");
//        cellStyle.setFontSize(10);
//        cellStyle.setColor("000000");
//        return cellStyle;
//    }
//
//    // 设置 table 格样式
//    public TableStyle getTableStyle() {
//        TableStyle tableStyle = new TableStyle();
//        tableStyle.setAlign(STJc.CENTER);
//        return tableStyle;
//    }
//
//    // 设置 cell 格样式
//    public Style getDataCellStyle() {
//        Style cellStyle = new Style();
//        cellStyle.setFontFamily("宋体");
//        cellStyle.setFontSize(8);
//        cellStyle.setColor("000000");
//        return cellStyle;
//    }
//
//    public VoteRuleSystem getSystem() {
//        return system;
//    }
//
//    public void setSystem(VoteRuleSystem system) {
//        this.system = system;
//    }
//
//    public int getDataSize() {
//        return dataSize;
//    }
//
//    public void setDataSize(int dataSize) {
//        this.dataSize = dataSize;
//    }
//}
