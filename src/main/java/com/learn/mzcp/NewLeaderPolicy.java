/*
 * Copyright (c) 2006- 2022, cuitianxin. All Rights Reserved.
 */

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
//import com.first.item.Item;
//import com.first.item.ItemDAO;
//import com.first.leaderfieldconfig.LeaderFieldConfig;
//import com.first.leaderfieldconfig.LeaderFieldConfigDAO;
//import com.first.voteelement.VoteElement;
//import com.first.voterule.VoteRuleSystem;
//import org.apache.poi.xwpf.usermodel.*;
//import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJc;
//
//import java.util.ArrayList;
//import java.util.LinkedHashMap;
//import java.util.List;
//
///**
// * @author ：Kristen
// * @date ：2022/10/10
// * @description : 新提拔的 word 封面逻辑生成策略
// */
//public class NewLeaderPolicy extends AbstractRenderPolicy<Object> {
//
//    private VoteElement element;
//    private VoteRuleSystem system;
//
//    private int dataSize;
//
//    // 所有的 voterule
//    //public ArrayList<VoteRule> voterTypeList;
//    public ArrayList<String> voterTypeList;
//    // 读取配置 list
//    public ArrayList<LeaderFieldConfig> configList;
//    // item（选人用人无一级指标）
//    public ArrayList<Item> itemList;
//    // 评分项
//    public String mobile_options;
//    // 评价问题
//    public String itemName;
//    // itemdic 的 itemId
//    private String itemId;
//
//    // 计算行和列
//    int col;
//    int row;
//    int base = 5;
//    LinkedHashMap<String, Integer> optionMap;
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
//
//        // 所有的 voterule
//        //voterTypeList = VoteRuleDAO.getVoteRuleListBySystem(system.getSystemID());
//        voterTypeList = this.constructVoterTypeList();
//        // 左侧第一列 item（没有一级指标的时候）
//        itemList = ItemDAO.getItemListByTable(element.getTableID());
//        // 获取评分项
//        mobile_options = itemList.get(0).getMobileOptions();
//        // 获取评价内容
//        itemName = itemList.get(0).getItemName();
//        // itemdic 的 itemId
//        itemId = itemList.get(0).getItemID();
//        // 将评分项的名称和分值放入 optionMap
//        optionMap = this.splitOption(mobile_options);
//        // 获取配置
//        String backSQL = "where menuIndex = '2' and usedelements='" + element.getElementCode() + "'";
//        configList = LeaderFieldConfigDAO.getLeaderFieldConfigListBySQL(backSQL);
//
//        // 计算行列
//        row = base + dataSize;
//        col = configList.size() + voterTypeList.size() * optionMap.size() + (optionMap.size() + 1) * 2 + 1;
//
//        // 当前位置插入表格
//        XWPFTable table = bodyContainer.insertNewTable(run, row, col);
//        this.setTableStyle(table);
//        this.setTableTitle(table);
//        this.setTableHeader(table);
//        this.setTableTag(table);
//        //this.setTableData(table);
//    }
//
//    // 整个 table 的样式在此设置
//    public void setTableStyle(XWPFTable table) {
//        // 设置 A4 幅面的平铺类型和列数
//        TableTools.widthTable(table, MiniTableRenderData.WIDTH_A4_NARROW_FULL, col);
//
//        // 设置 border
//        TableTools.borderTable(table, 10);
//        for (XWPFTableRow tableRow : table.getRows()) {
//            for (int i = 0; i < tableRow.getTableCells().size(); i++) {
//                tableRow.getCell(i).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
//                tableRow.getCell(i).setWidth("500");
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
//        RowRenderData header0 = RowRenderData.build(new TextRenderData(title, cellStyle));
//        header0.setRowStyle(tableStyle);
//        TableTools.mergeCellsHorizonal(table, 0, 0, col - 1);
//        MiniTableRenderPolicy.Helper.renderRow(table, 0, header0);
//    }
//
//    public void setTableHeader(XWPFTable table) {
//        // 第1行值的数组
//        String[] strHeader1 = new String[2];
//        strHeader1[0] = "被评议对象的基本情况";
//        strHeader1[1] = itemName;
//
//        // 第1行水平合并单元格 对提拔任用该干部的看法部分
//        TableTools.mergeCellsHorizonal(table, 1, configList.size() + 1, col - 1);
//
//        // 第1行水平合并单元格 被评议对象的基本情况部分
//        TableTools.mergeCellsHorizonal(table, 1, 0, configList.size());
//        TableTools.mergeCellsHorizonal(table, 2, 0, configList.size());
//
//        // 第1行垂直合并单元格
//        TableTools.mergeCellsVertically(table, 0, 1, 2);
//
//        // 第2行值的数组
//        String[] strHeader2 = new String[voterTypeList.size() + 2];
//        for (int i = 0; i < voterTypeList.size(); i++) {
//            //strHeader2[i + 1] = voterTypeList.get(i).getShortName();
//            strHeader2[i + 1] = voterTypeList.get(i).replaceAll(";", "_");
//        }
//        strHeader2[strHeader2.length - 1] = "合计";
//
//        // 第2行水平合并单元格 票种部分
//        int startHeader2 = 1;
//        int endHeader2;
//        for (int i = 0; i < strHeader2.length - 2; i++) {
//            endHeader2 = startHeader2 + optionMap.size() - 1;
//            TableTools.mergeCellsHorizonal(table, 2, startHeader2, endHeader2);
//            startHeader2++;
//        }
//
//        // 第2行水平合并单元格 合计部分
//        TableTools.mergeCellsHorizonal(table, 2, startHeader2, startHeader2 + 2 * (optionMap.size() + 1) - 1);
//
//        // 第3行值的数组
//        String[] strHeader3 = new String[configList.size() + (optionMap.size()) * (voterTypeList.size() + 1) + 2];
//        strHeader3[0] = "序号";
//        for (int i = 0; i < configList.size(); i++) {
//            strHeader3[i + 1] = configList.get(i).getFieldName();
//        }
//        int header3Index = configList.size() + 1;
//        for (int i = 0; i < voterTypeList.size() + 1; i++) {
//            for (int j = 0; j < optionMap.keySet().toArray().length; j++) {
//                strHeader3[header3Index++] = optionMap.keySet().toArray()[j].toString();
//            }
//        }
//        strHeader3[strHeader3.length - 1] = "认同度";
//
//        // 第3行水平合并单元格 合计部分
//        int startHeader3 = voterTypeList.size() * optionMap.size() + configList.size() + 1;
//        for (int i = 0; i < optionMap.size(); i++) {
//            TableTools.mergeCellsHorizonal(table, 3, startHeader3, startHeader3 + 1);
//            startHeader3++;
//        }
//
//        // 第3行水平合并单元格 认同度部分
//        TableTools.mergeCellsHorizonal(table, 3, startHeader3, startHeader3 + 1);
//
//        // 第3行垂直合并单元格 属性部分
//        for (int i = 0; i < configList.size() + 1; i++) {
//            TableTools.mergeCellsVertically(table, i, 3, 4);
//        }
//
//        // 第4行数组的值
//        String[] strHeader4 = new String[col];
//        for (int i = 0; i < voterTypeList.size() * optionMap.size(); i++) {
//            strHeader4[i + configList.size() + 1] = "得票";
//        }
//
//        // 第4行值的数组 合计部分
//        for (int i = voterTypeList.size() * optionMap.size() + configList.size() + 1; i < strHeader4.length; i++) {
//            if ((voterTypeList.size() * optionMap.size() + configList.size() + 1) % 2 == 0) {
//                if (i % 2 == 0) {
//                    strHeader4[i] = "得票";
//                } else {
//                    strHeader4[i] = "比例";
//                }
//            } else {
//                if (i % 2 == 0) {
//                    strHeader4[i] = "比例";
//                } else {
//                    strHeader4[i] = "得票";
//                }
//            }
//        }
//        strHeader4[strHeader4.length - 1] = "排名";
//
//        // 构建第 1-4 行
//        Style style = this.getCellStyle();
//        RowRenderData header1 = this.build(strHeader1, style);
//        RowRenderData header2 = this.build(strHeader2, style);
//        RowRenderData header3 = this.build(strHeader3, style);
//        RowRenderData header4 = this.build(strHeader4, style);
//        TableStyle tableStyle = this.getTableStyle();
//        header1.setRowStyle(tableStyle);
//        header2.setRowStyle(tableStyle);
//        header3.setRowStyle(tableStyle);
//        header4.setRowStyle(tableStyle);
//        MiniTableRenderPolicy.Helper.renderRow(table, 1, header1);
//        MiniTableRenderPolicy.Helper.renderRow(table, 2, header2);
//        MiniTableRenderPolicy.Helper.renderRow(table, 3, header3);
//        MiniTableRenderPolicy.Helper.renderRow(table, 4, header4);
//    }
//
//    // 设置标签
//    public void setTableTag(XWPFTable table) {
//        String[] strTag = new String[col];
//        for (int i = 0; i < dataSize; i++) {
//            // 设置 tag（人员基本情况部分）
//            strTag[0] = "{{sequence_" + i + "}}";
//            for (int j = 0; j < configList.size(); j++) {
//                strTag[j + 1] = "{{" + configList.get(j).getFieldID() + "_" + i + "}}";
//            }
//
//            // 设置 tag（票种部分）
//            int index = configList.size() + 1;
////            for (VoteRule voteRule : voterTypeList) {
////                String voteType = voteRule.getShortName();
////                for (int k = 0; k < optionMap.size(); k++) {
////                    // strTag[index] = "{{count_leader21_" + optionMap.values().toArray()[k].toString() + "__" + voteType.replaceAll("、", "_") + "_" + i + "}}";
////                    strTag[index] = "{{count_" + itemId + "_" + optionMap.values().toArray()[k].toString() + "__" + voteType.replaceAll("、", "_") + "_" + i + "}}";
////                    index++;
////                }
////            }
//
//            for (String voteType : voterTypeList) {
//                for (int k = 0; k < optionMap.size(); k++) {
//                    // strTag[index] = "{{count_leader21_" + optionMap.values().toArray()[k].toString() + "__" + voteType.replaceAll("、", "_") + "_" + i + "}}";
//                    strTag[index] = "{{count_" + itemId + "_" + optionMap.values().toArray()[k].toString() + "__" + voteType.replaceAll(";", "_") + "_" + i + "}}";
//                    index++;
//                }
//            }
//
//            // 设置 tag（合计部分）
//            for (int j = 0; j < optionMap.values().toArray().length; j++) {
//                strTag[index++] = "{{count_" + itemId + "_" + optionMap.values().toArray()[j].toString() + "_" + i + "}}";
//                strTag[index++] = "{{rate_" + itemId + "_" + optionMap.values().toArray()[j].toString() + "_" + i + "}}";
//            }
//
//            // 设置 tag（最后一行部分）
//            strTag[col - 2] = "{{rate_" + itemId + "_7_" + i + "}}";
//            strTag[col - 1] = "{{sort_" + (col - 2) + "_" + i + "}}";
//
//            // 构建
//            Style style = this.getCellStyle();
//            RowRenderData tag = this.build(strTag, style);
//            TableStyle tableStyle = this.getTableStyle();
//            tag.setRowStyle(tableStyle);
//            MiniTableRenderPolicy.Helper.renderRow(table, i + base, tag);
//        }
//    }
//
//    // 将题目的选项 如：“4:不了解:0;6:不认同:0;8:基本认同:0;10:认同:0” 存入 map，key 为显示的值，value 为分值
//    public LinkedHashMap<String, Integer> splitOption(String option) {
//        LinkedHashMap<String, Integer> optionMap = new LinkedHashMap<>();
//        String[] strArray = option.replaceAll(":0", "").split(";");
//        for (int i = strArray.length - 1; i >= 0; i--) {
//            String[] str = strArray[i].split(":");
//            optionMap.put(str[1], Integer.valueOf(str[0]));
//        }
//        return optionMap;
//    }
//
//    // 分票
//    public ArrayList<String> constructVoterTypeList() {
//        ArrayList<String> list = new ArrayList<>();
//        list.add("A1;A2");
//        list.add("A3");
//        list.add("B");
//        list.add("C");
//        return list;
//    }
//
//    // 根据 String[] 构建一行的数据，同一行使用一个 Style
//    public RowRenderData build(String[] cellStr, Style style) {
//        List<TextRenderData> data = new ArrayList<>();
//        if (null != cellStr) {
//            for (String col : cellStr) {
//                data.add(new TextRenderData(col, style));
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
//    public VoteElement getElement() {
//        return element;
//    }
//
//    public void setElement(VoteElement element) {
//        this.element = element;
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
