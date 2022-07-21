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
//import com.first.group.ItemGroup;
//import com.first.item.Item;
//import com.first.voterule.VoteRule;
//import org.apache.poi.xwpf.usermodel.*;
//import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJc;
//
//import java.util.ArrayList;
//import java.util.HashMap;
//import java.util.LinkedHashMap;
//import java.util.List;
//
///**
// * @author ：Kristen
// * @date ：2022/6/22
// * @description :
// */
//public class GroupPolicy extends AbstractRenderPolicy<Object> {
//
//    // 左侧第一列 group
//    private ArrayList<ItemGroup> groupList = new ArrayList<>();
//    // 左侧第二列 items（有一级指标的时候）
//    private ArrayList<Item> itemsList = new ArrayList<>();
//    // 左侧第一列 item（没有一级指标的时候）
//    private ArrayList<Item> itemList = new ArrayList<>();
//    // 所有的 voterule
//    private ArrayList<VoteRule> voterTypeList = new ArrayList<>();
//    // 第二行 voterule 的所有分组
//    private ArrayList<String> voteTypeGroupList = new ArrayList<>();
//    // 第三行 voterule 分组的子票种
//    private ArrayList<ArrayList<VoteRule>> innerTitleList = new ArrayList<>();
//    // 第二行分组之外的 voterule
//    private ArrayList<VoteRule> outerTitleList = new ArrayList<>();
//    private LinkedHashMap<String, ArrayList<VoteRule>> voteGroupMap = new LinkedHashMap<>();
//    private String year;
//    private String organCode;
//    private String wsID;
//
//    // 计算行
//    int row;
//    // 计算列、先判断有无外部董事
//    boolean isHaveWBDS;
//    int col;
//
//    @Override
//    protected void afterRender(RenderContext<Object> renderContext) {
//        // 清空标签
//        clearPlaceholder(renderContext, true);
//    }
////
////    @Override
//    public void doRender(RenderContext<Object> renderContext) throws Exception {
//        XWPFRun run = renderContext.getRun();
//        // 当前位置的容器
//        BodyContainer bodyContainer = BodyContainerFactory.getBodyContainer(run);
//        voteTypeGroupList.addAll(voteGroupMap.keySet());
//        for (String s : voteTypeGroupList) {
//            innerTitleList.add(voteGroupMap.get(s));
//        }
//        row = this.getItemsList().size() + 4;
//        isHaveWBDS = this.checkWBDS(voterTypeList);
//        col = this.calculateColumn(voterTypeList, voteTypeGroupList, isHaveWBDS);
//
//        // 当前位置插入表格
//        XWPFTable table = bodyContainer.insertNewTable(run, row, col);
//        this.setTableStyle(table);
//        this.setTableTitle(table);
//        this.setTableHeader(table);
//        this.setGroupAndItem(table);
//        this.setCellTag(table);
//        this.setLastRow(table);
//    }
//
//    // 整个 table 的样式在此设置
//    public void setTableStyle(XWPFTable table) {
//        // 设置 A4 幅面的平铺类型和列数
//        if (col <= 15) {
//            TableTools.widthTable(table, MiniTableRenderData.WIDTH_A4_FULL, col);
//        } else {
//            TableTools.widthTable(table, MiniTableRenderData.WIDTH_A4_NARROW_FULL, col);
//        }
//        TableTools.borderTable(table, 10);
//        for (XWPFTableRow tableRow : table.getRows()) {
//            for (int i = 0; i < tableRow.getTableCells().size(); i++) {
//                tableRow.getCell(i).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
//                tableRow.getCell(i).setWidth("500");
//            }
//        }
//        table.setCellMargins(5, 10, 5, 10);
//        table.setTableAlignment(TableRowAlign.CENTER);
//    }
//
//    // 设置第一行标题的样式
//    public void setTableTitle(XWPFTable table) {
//        Style cellStyle = new Style();
//        cellStyle.setFontSize(12);
//        cellStyle.setColor("000000");
//        cellStyle.setFontFamily("黑体");
//        TableStyle style = new TableStyle();
//        style.setAlign(STJc.CENTER);
//        style.setBackgroundColor("DCDCDC");
//        String title = organCode + "领导班子" + year + "年度综合测评汇总表";
//        RowRenderData header0 = RowRenderData.build(new TextRenderData(title, cellStyle));
//        header0.setRowStyle(style);
//        TableTools.mergeCellsHorizonal(table, 0, 0, col - 1);
//        MiniTableRenderPolicy.Helper.renderRow(table, 0, header0);
//    }
//
//    // 设置第二、三行标题的样式
//    public void setTableHeader(XWPFTable table) {
//        // 含有分组
//        if ((voteTypeGroupList != null && voteTypeGroupList.size() > 0) && (innerTitleList != null && innerTitleList.size() > 0)) {
//            // 处理第二行，先判断第二行的表头内容
//            int length = voteTypeGroupList.size() + 2;
//            // 判断是否包含外部董事
//            if (isHaveWBDS) {
//                length++;
//            }
//            String[] strHeader1 = new String[length];
//            strHeader1[0] = "评价内容";
//            strHeader1[length - 1] = "全体";
//            if (isHaveWBDS) {
//                for (VoteRule rule : voterTypeList) {
//                    if (rule.getFilterName().contains("外部董事")) {
//                        strHeader1[length - 2] = rule.getFilterName() + "（" + rule.getRuleName() + "）";
//                    }
//                }
//            }
//            for (int i = 0; i < voteTypeGroupList.size(); i++) {
//                strHeader1[i + 1] = voteTypeGroupList.get(i);
//            }
//            // 构建第二行
//            Style cellStyle = this.getCellStyle();
//            RowRenderData header1 = this.build(cellStyle, strHeader1);
//            // 垂直合并 “全体”
//            TableTools.mergeCellsVertically(table, col - 1, 1, 2);
//            TableTools.mergeCellsVertically(table, col - 2, 1, 2);
//
//            // 是否含有外部董事
//            if (isHaveWBDS) {
//                // 垂直合并 “外部董事”
//                TableTools.mergeCellsVertically(table, col - 4, 1, 2);
//            }
//            // 垂直合并“评价内容”
//            TableTools.mergeCellsVertically(table, 0, 1, 2);
//
//            // 水平合并“评价内容”
//            TableTools.mergeCellsHorizonal(table, 1, 0, 1);
//            TableTools.mergeCellsHorizonal(table, 2, 0, 1);
//
//            // 水平合并“分组”
//            int start = 1;
//            int end;
//            for (ArrayList<VoteRule> rules : innerTitleList) {
//                end = (rules.size() + 1) * 2;
//                TableTools.mergeCellsHorizonal(table, 1, start, start + end - 1);
//                start++;
//            }
//
//            // 判断是否含有外部董事
//            if (isHaveWBDS) {
//                // 水平合并“外部董事”
//                TableTools.mergeCellsHorizonal(table, 1, voteTypeGroupList.size() + 1, voteTypeGroupList.size() + 2);
//                TableTools.mergeCellsHorizonal(table, 2, col - voteTypeGroupList.size() - 4, col - voteTypeGroupList.size() - 3);
//                // 水平合并“全体”
//                TableTools.mergeCellsHorizonal(table, 1, voteTypeGroupList.size() + 2, voteTypeGroupList.size() + 3);
//                TableTools.mergeCellsHorizonal(table, 2, col - voteTypeGroupList.size() - 3, col - voteTypeGroupList.size() - 2);
//            } else {
//                TableTools.mergeCellsHorizonal(table, 1, voteTypeGroupList.size() + 1, voteTypeGroupList.size() + 2);
//                TableTools.mergeCellsHorizonal(table, 2, col - voteTypeGroupList.size() - 1, col - voteTypeGroupList.size());
//            }
//            TableStyle tableStyle = this.getTableStyle();
//            header1.setRowStyle(tableStyle);
//            MiniTableRenderPolicy.Helper.renderRow(table, 1, header1);
//
//            // 处理分组内的票种和小计
//            int total = countRow(innerTitleList);
//            int lengthItem = (total + voteTypeGroupList.size()) + 1;
//            String[] strHeader2 = new String[lengthItem];
//            int index = 1;
//            for (ArrayList<VoteRule> rules : innerTitleList) {
//                for (VoteRule rule : rules) {
//                    strHeader2[index++] = rule.getFilterName() + "（" + rule.getRuleName() + "）";
//                }
//                strHeader2[index++] = "小计";
//            }
//            RowRenderData header2 = this.build(cellStyle, strHeader2);
//            // 分别给以上几个 TextRenderData 设置合并单元格、第一个格子必须写空值否则无法写入到 table
//            for (int i = 0; i < strHeader2.length - 1; i++) {
//                TableTools.mergeCellsHorizonal(table, 2, i + 1, i + 2);
//            }
//            header2.setRowStyle(tableStyle);
//            MiniTableRenderPolicy.Helper.renderRow(table, 2, header2);
//        } else {
//            int length = voterTypeList.size() + 2;
//            String[] strHeader1 = new String[length];
//            strHeader1[0] = "评价内容";
//            strHeader1[length - 1] = "全体";
//            for (int i = 0; i < voterTypeList.size(); i++) {
//                strHeader1[i + 1] = voterTypeList.get(i).getFilterName();
//            }
//            // 构建第二行
//            Style cellStyle = this.getCellStyle();
//            RowRenderData header1 = this.build(cellStyle, strHeader1);
//            for (int i = 0; i < voterTypeList.size() + 2; i++) {
//                int j = i + 1;
//                TableTools.mergeCellsHorizonal(table, 1, i, j);
//                TableTools.mergeCellsHorizonal(table, 2, i, j);
//            }
//
//            for (int i = voterTypeList.size() + 1; i >= 0; i--) {
//                TableTools.mergeCellsVertically(table, i, 1, 2);
//            }
//            TableStyle tableStyle = this.getTableStyle();
//            header1.setRowStyle(tableStyle);
//            MiniTableRenderPolicy.Helper.renderRow(table, 1, header1);
//        }
//    }
//
//    // group 的设置、集中在第1列
//    public void setGroupAndItem(XWPFTable table) {
//        TableStyle tableStyle = this.getTableStyle();
//        Style cellStyle = this.getCellStyle();
//        String[] str = new String[col];
//        HashMap<String, ItemGroup> groupMap = new HashMap<>();
//        for (ItemGroup group : groupList) {
//            groupMap.put(group.getGroupID(), group);
//        }
//        for (int i = 0; i < itemsList.size(); i++) {
//            Item item = itemsList.get(i);
//            str[0] = groupMap.get(item.getGroupID()).getGroupName();
//            str[1] = item.getItemName();
//            RowRenderData rowData = this.build(cellStyle, str);
//            rowData.setRowStyle(tableStyle);
//            MiniTableRenderPolicy.Helper.renderRow(table, i + 3, rowData);
//        }
//        HashMap<String, Integer> map = new HashMap<>();
//        for (Item item : itemsList) {
//            if (!map.containsKey(item.getGroupID())) {
//                map.put(item.getGroupID(), 0);
//            }
//            map.put(item.getGroupID(), map.get(item.getGroupID()) + 1);
//        }
//
//        int startRowIndex = 3;
//        for (ItemGroup group : groupList) {
//            int rowCount = map.get(group.getGroupID());
//            TableTools.mergeCellsVertically(table, 0, startRowIndex, startRowIndex + rowCount - 1);
//            startRowIndex += rowCount;
//        }
//    }
//
//    // 设置 table 的 cell 中的 tag
//    public void setCellTag(XWPFTable table) {
//        String[] str = new String[col];
//        ArrayList<VoteRule> list = getAllRules();
//        HashMap<String, ItemGroup> groupMap = new HashMap<>();
//        for (ItemGroup group : groupList) {
//            groupMap.put(group.getGroupID(), group);
//        }
//        for (int i = 0; i < itemsList.size(); i++) {
//            Item item = itemsList.get(i);
//            for (int j = 2; j < str.length; j += 2) {
//                if (j < str.length - 2) {
//                    if (list.get((j - 2) / 2).getVoterType().contains("rule")) {
//                        str[j] = "{{all100_" + item.getItemID() + "_" + list.get((j - 2) / 2).getVoterType().replaceAll(";", "") + "}}";
//                        str[j + 1] = "{{all100_" + groupMap.get(item.getGroupID()).getGroupID() + "_" + list.get((j - 2) / 2).getVoterType().replaceAll(";", "") + "}}";
//                    } else {
//                        str[j] = "{{avg_" + item.getItemID() + "_" + list.get((j - 2) / 2).getVoterType().replaceAll(";", "") + "}}";
//                        str[j + 1] = "{{avg_" + groupMap.get(item.getGroupID()).getGroupID() + "_" + list.get((j - 2) / 2).getVoterType().replaceAll(";", "") + "}}";
//                    }
//                } else {
//                    if (j == str.length - 2) {
//                        str[j] = "{{avg_" + item.getItemID() + "}}";
//                        str[j + 1] = "{{avg_" + groupMap.get(item.getGroupID()).getGroupID() + "}}";
//                    }
//                }
//            }
//            Style cellStyle = this.getDataCellStyle();
//            RowRenderData rowData = this.build(cellStyle, str);
//            TableStyle tableStyle = this.getTableStyle();
//            rowData.setRowStyle(tableStyle);
//            MiniTableRenderPolicy.Helper.renderRow(table, i + 3, rowData);
//        }
//        HashMap<String, Integer> map = new HashMap<>();
//        for (Item item : itemsList) {
//            if (!map.containsKey(item.getGroupID())) {
//                map.put(item.getGroupID(), 0);
//            }
//            map.put(item.getGroupID(), map.get(item.getGroupID()) + 1);
//        }
//
//        int startRowIndex = 3;
//        for (ItemGroup group : groupList) {
//            int rowCount = map.get(group.getGroupID());
//            int column = 3;
//            for (int i = 0; i < col / 2 - 1; i++) {
//                TableTools.mergeCellsVertically(table, column, startRowIndex, startRowIndex + rowCount - 1);
//                column += 2;
//            }
//            startRowIndex += rowCount;
//        }
//    }
//
//    // 设置最后一行
//    public void setLastRow(XWPFTable table) {
//        // 最后一行的加权总得分
//        String[] str = new String[col / 2];
//        str[0] = "加权汇总得分";
//        ArrayList<VoteRule> list = getAllRules();
//        for (int i = 0; i < list.size(); i++) {
//            if (list.get(i).getVoterType().contains("rule")) {
//                str[i + 1] = "{{all100_" + list.get(i).getVoterType().replaceAll(";", "") + "}}";
//            } else {
//                str[i + 1] = "{{avg_" + list.get(i).getVoterType().replaceAll(";", "") + "}}";
//            }
//        }
//        str[str.length - 1] = "{{avg}}";
//        Style cellStyle = this.getCellStyle();
//        RowRenderData total = this.build(cellStyle, str);
//        for (int i = 0; i < col / 2; i++) {
//            TableTools.mergeCellsHorizonal(table, row - 1, i, i + 1);
//        }
//        TableStyle tableStyle = this.getTableStyle();
//        total.setRowStyle(tableStyle);
//        MiniTableRenderPolicy.Helper.renderRow(table, row - 1, total);
//    }
//
//    // 判断否含有外部董事
//    public boolean checkWBDS(List<VoteRule> voterTypeList) {
//        for (VoteRule rule : voterTypeList) {
//            if (rule.getFilterName().contains("外部董事")) {
//                isHaveWBDS = true;
//                break;
//            }
//        }
//        return isHaveWBDS;
//    }
//
//    // 计算列数
//    public int calculateColumn(List<VoteRule> voterTypeList, List<String> voteTypeGroupList, boolean isHaveWBDS) {
//        int column;
//        int littleCount = 0;
//        // 如果分组大于 1，则无论有无外部董事，每个组都有一个小计
//        if (voteTypeGroupList.size() > 1) {
//            littleCount = voteTypeGroupList.size() * 2;
//        } else {
//            // 只有一个分组的情况下，如果有外部董事则有一个小计，如果没有外部董事则没有小计
//            if (isHaveWBDS) {
//                littleCount += 2;
//            }
//        }
//        column = voterTypeList.size() * 2 + littleCount + 4;
//        return column;
//    }
//
//    // 计算所有分组的项的个数
//    public int countRow(ArrayList<ArrayList<VoteRule>> innerTitleList) {
//        int count = 0;
//        for (ArrayList<VoteRule> rules : innerTitleList) {
//            for (int j = 0; j < rules.size(); j++) {
//                count++;
//            }
//        }
//        return count;
//    }
//
//    // 拼接所有的 votetype
//    public ArrayList<VoteRule> getAllRules() {
//        ArrayList<VoteRule> allRules = new ArrayList<>();
//        VoteRule voteRule = new VoteRule();
//        // 将分组内的票种和小计的票种放入 allRules
//        for (ArrayList<VoteRule> voteRules : innerTitleList) {
//            allRules.addAll(voteRules);
//            StringBuilder type = new StringBuilder();
//            for (VoteRule rule : voteRules) {
//                type.append(rule.getRuleID()).append(";");
//            }
//            voteRule.setVoterType(type.substring(0, type.length() - 1));
//            allRules.add(voteRule);
//        }
//        // 将 voterTypeList 中已经在 innerTitleList 的票种去除
//        for (VoteRule allRule : allRules) {
//            for (int i = 0; i < voterTypeList.size(); i++) {
//                if (voterTypeList.get(i).getRuleName().equals(allRule.getRuleName())) {
//                    voterTypeList.remove(i);
//                }
//            }
//        }
//        allRules.addAll(voterTypeList);
//        return allRules;
//    }
//
//    // 根据 String[] 构建一行的数据，同一行使用一个 Style
//    public RowRenderData build(Style style, String... cellStr) {
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
//    public ArrayList<ItemGroup> getGroupList() {
//        return groupList;
//    }
//
//    public void setGroupList(ArrayList<ItemGroup> groupList) {
//        this.groupList = groupList;
//    }
//
//    public ArrayList<Item> getItemsList() {
//        return itemsList;
//    }
//
//    public void setItemsList(ArrayList<Item> itemsList) {
//        this.itemsList = itemsList;
//    }
//
//    public ArrayList<Item> getItemList() {
//        return itemList;
//    }
//
//    public void setItemList(ArrayList<Item> itemList) {
//        this.itemList = itemList;
//    }
//
//    public ArrayList<VoteRule> getVoterTypeList() {
//        return voterTypeList;
//    }
//
//    public void setVoterTypeList(ArrayList<VoteRule> voterTypeList) {
//        this.voterTypeList = voterTypeList;
//    }
//
//    public ArrayList<String> getVoteTypeGroupList() {
//        return voteTypeGroupList;
//    }
//
//    public void setVoteTypeGroupList(ArrayList<String> voteTypeGroupList) {
//        this.voteTypeGroupList = voteTypeGroupList;
//    }
//
//    public LinkedHashMap<String, ArrayList<VoteRule>> getVoteGroupMap() {
//        return voteGroupMap;
//    }
//
//    public void setVoteGroupMap(LinkedHashMap<String, ArrayList<VoteRule>> voteGroupMap) {
//        this.voteGroupMap = voteGroupMap;
//    }
//
//    public ArrayList<ArrayList<VoteRule>> getInnerTitleList() {
//        return innerTitleList;
//    }
//
//    public void setInnerTitleList(ArrayList<ArrayList<VoteRule>> innerTitleList) {
//        this.innerTitleList = innerTitleList;
//    }
//
//    public ArrayList<VoteRule> getOuterTitleList() {
//        return outerTitleList;
//    }
//
//    public void setOuterTitleList(ArrayList<VoteRule> outerTitleList) {
//        this.outerTitleList = outerTitleList;
//    }
//
//    public String getYear() {
//        return year;
//    }
//
//    public void setYear(String year) {
//        this.year = year;
//    }
//
//    public String getOrganCode() {
//        return organCode;
//    }
//
//    public void setOrganCode(String organCode) {
//        this.organCode = organCode;
//    }
//
//    public String getWsID() {
//        return wsID;
//    }
//
//    public void setWsID(String wsID) {
//        this.wsID = wsID;
//    }
//}