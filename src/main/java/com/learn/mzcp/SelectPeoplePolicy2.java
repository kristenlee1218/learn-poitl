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
// * @description : 选人用人表 2 的 word 封面逻辑生成策略
// */
//public class SelectPeoplePolicy2 extends AbstractRenderPolicy<Object> {
//
//    private VoteElement element;
//    private VoteRuleSystem system;
//
//    // 左侧第一列 item（选人用人无一级指标）
//    public ArrayList<Item> itemList;
//    // 所有的 VoteRule
//    //public ArrayList<VoteRule> voterTypeList;
//    public ArrayList<String> voterTypeList;
//    // 评分项
//    public String mobile_options;
//    // 评价问题
//    public String itemName;
//
//    // 计算行和列
//    int col;
//    int row;
//    int base = 4;
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
//        String backSQL = "and isvote = 'Y' and ischeckbox = 'Y' and mobile_mode = '2'";
//        itemList = ItemDAO.getMobileItemList(element.getTableID(), backSQL);
//        // 获取评分项
//        mobile_options = itemList.get(0).getMobileOptions();
//        // 获取评价内容
//        itemName = itemList.get(0).getItemName();
//        // 将评分项的名称和分值放入 optionMap
//        optionMap = this.splitOption(mobile_options);
//
//        // 计算行列
//        col = (voterTypeList.size() + 1) * 2 + 1;
//        row = optionMap.size() + base;
//
//        // 当前位置插入表格
//        XWPFTable table = bodyContainer.insertNewTable(run, row, col);
//        this.setTableStyle(table);
//        this.setTableTitle(table);
//        this.setTableQuestion(table);
//        this.setTableHeader(table);
//        this.setTableItem(table);
//        this.setTableTag(table);
//        // this.setTableData(table);
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
//    // 设置第一行标题的样式 {{title}}
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
//    // 设置第二行问题的样式
//    public void setTableQuestion(XWPFTable table) {
//        Style cellStyle = new Style();
//        cellStyle.setFontSize(8);
//        cellStyle.setColor("000000");
//        cellStyle.setFontFamily("黑体");
//        TableStyle tableStyle = new TableStyle();
//        RowRenderData header1 = RowRenderData.build(new TextRenderData(itemName, cellStyle));
//        header1.setRowStyle(tableStyle);
//        TableTools.mergeCellsHorizonal(table, 1, 0, col - 1);
//        MiniTableRenderPolicy.Helper.renderRow(table, 1, header1);
//    }
//
//    public void setTableHeader(XWPFTable table) {
//        // 第2行值的数组
//        String[] strHeader1 = new String[voterTypeList.size() + 2];
//        strHeader1[0] = "结果";
//        for (int i = 0; i < voterTypeList.size(); i++) {
//            //strHeader1[i + 1] = voterTypeList.get(i).getShortName();
//            strHeader1[i + 1] = voterTypeList.get(i).replaceAll(";","_");
//        }
//        strHeader1[strHeader1.length - 1] = "合计";
//
//        // 第2行垂直合并单元格
//        TableTools.mergeCellsVertically(table, 0, 2, 3);
//
//        // 第2行水平合并单元格
//        int startHeader1 = 1;
//        for (int i = 0; i < voterTypeList.size() + 1; i++) {
//            TableTools.mergeCellsHorizonal(table, 2, startHeader1, startHeader1 + 1);
//            startHeader1++;
//        }
//
//        // 第3行值的数组 票种部分
//        String[] strHeader2 = new String[col];
//        for (int i = 0; i < (voterTypeList.size() + 1) * 2; i++) {
//            if (i % 2 == 1) {
//                strHeader2[i + 1] = "票数";
//            } else {
//                strHeader2[i + 1] = "百分百";
//            }
//        }
//
//        // 构建第2-3行
//        Style style = this.getCellStyle();
//        RowRenderData header1 = this.build(strHeader1, style);
//        RowRenderData header2 = this.build(strHeader2, style);
//        TableStyle tableStyle = this.getTableStyle();
//        header1.setRowStyle(tableStyle);
//        header2.setRowStyle(tableStyle);
//        MiniTableRenderPolicy.Helper.renderRow(table, 2, header1);
//        MiniTableRenderPolicy.Helper.renderRow(table, 3, header2);
//    }
//
//    // 设置第1列
//    public void setTableItem(XWPFTable table) {
//        Style cellDataStyle = this.getDataCellStyle();
//        TableStyle tableStyle = this.getTableStyle();
//        for (int i = 0; i < optionMap.size(); i++) {
//            RowRenderData itemData = RowRenderData.build(new TextRenderData((i + 1) + "、" + optionMap.keySet().toArray()[i].toString(), cellDataStyle));
//            itemData.setRowStyle(tableStyle);
//            MiniTableRenderPolicy.Helper.renderRow(table, i + base, itemData);
//        }
//    }
//
//    // 设置标签
//    public void setTableTag(XWPFTable table) {
//        for (int i = base; i < row; i++) {
//            String[] strTag = new String[col];
//
//            // 设置 tag（票种部分）
//            int index = 1;
//            for (String s : voterTypeList) {
//                //String voteType = voterTypeList.get(i).getShortName();
//                strTag[index++] = "{{sect#CONCAT(" + itemList.get(0).getItemID() + ",'') like '%" + (i - base + 1) + "%'#@" + s.replaceAll(";", "") + "}}";
//                strTag[index++] = "{{sectrate#CONCAT(" + itemList.get(0).getItemID() + ",'') like '%" + (i - base + 1) + "%'#@" + s.replaceAll(";", "") + "}}";
//                strTag[col - 2] = "{{sect#CONCAT(" + itemList.get(0).getItemID() + ",'') like '%" + (i - base + 1) + "%'#}}";
//                strTag[col - 1] = "{{sectrate#CONCAT(" + itemList.get(0).getItemID() + ",'') like '%" + (i - base + 1) + "%'#}}";
//            }
//            Style style = this.getCellStyle();
//            RowRenderData tag = this.build(strTag, style);
//            TableStyle tableStyle = this.getTableStyle();
//            tag.setRowStyle(tableStyle);
//            MiniTableRenderPolicy.Helper.renderRow(table, i, tag);
//        }
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
//    // 将题目的选项 如：“1:落实党中央关于领导班子和干部队伍建设工作要求有差距:0;
//    // 2:选人用人把关不严、质量不高:0;3:坚持事业为上不够，不能做到以事择人、人岗相适:0;
//    // 4:激励担当作为用人导向不鲜明，论资排辈情况严重:0;5:选人用人“个人说了算”:0;
//    // 6:任人唯亲、拉帮结派:0;7:跑官要官、买官卖官、说情打招呼:0;
//    // 8:执行干部选拔任用政策规定不严格:0;9:干部队伍建设统筹谋划不够，结构不合理:0;
//    // 10:干部队伍能力素质不适应工作要求:0” 存入 map，key 为显示的值，value 为分值
//    public LinkedHashMap<String, Integer> splitOption(String option) {
//        LinkedHashMap<String, Integer> optionMap = new LinkedHashMap<>();
//        String[] strArray = option.replaceAll(":0", "").split(";");
//        for (String s : strArray) {
//            String[] str = s.split(":");
//            optionMap.put(str[1], Integer.valueOf(str[0]));
//        }
//        return optionMap;
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
//}
