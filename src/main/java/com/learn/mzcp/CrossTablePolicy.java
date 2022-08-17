///*
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
//import com.first.group.ItemGroupDAO;
//import com.first.item.Item;
//import com.first.item.ItemDAO;
//import com.first.voteelement.VoteElement;
//import com.first.voterule.VoteRule;
//import com.first.voterule.VoteRuleDAO;
//import com.first.voterule.VoteRuleSystem;
//import com.waf.db.DAOException;
//import org.apache.poi.xwpf.usermodel.*;
//import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJc;
//
//import java.util.*;
//
//*/
///**
// * @author ：Kristen
// * @date ：2022/6/22
// * @description :
// *//*
//
//public class CrossTablePolicy extends AbstractRenderPolicy<Object> {
//
//    private VoteElement element;
//    private VoteRuleSystem system;
//
//    // 左侧第一列 group
//    public ArrayList<ItemGroup> groupList;
//    // 左侧第二列 items（有一级指标的时候）
//    public ArrayList<Item> itemsList;
//    // 左侧第一列 item（没有一级指标的时候）
//    public ArrayList<Item> itemList;
//    // 所有的 voterule
//    public ArrayList<VoteRule> voterTypeList;
//    // 第二行 voterule 的所有分组
//    public ArrayList<String> voteTypeGroupList = new ArrayList<>();
//    // 第三行 voterule 分组的子票种
//    public ArrayList<ArrayList<VoteRule>> innerTitleList = new ArrayList<>();
//    // 第二行分组之外的 voterule
//    public LinkedHashMap<String, ArrayList<VoteRule>> voteGroupMap;
//
//    // 计算行
//    int row;
//    // 计算列、受限判断是否包含一级指标、如果包含则先判断有无外部董事
//    boolean isHaveGroup;
//    // 是否含有外部分组
//    boolean isHaveWBDS;
//    int col;
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
//        // 左侧第一列 group
//        groupList = ItemGroupDAO.getGroupListByTable(element.getTableID());
//        // 左侧第二列 items（有一级指标的时候）
//        itemsList = ItemDAO.getItemListByTable(element.getTableID());
//        // 左侧第一列 item（没有一级指标的时候）
//        itemList = ItemDAO.getItemListByTable(element.getTableID());
//        // 所有的 voterule
//        voterTypeList = VoteRuleDAO.getVoteRuleListBySystem(system.getSystemID());
//
//        // 组名和分组对应
//        voteGroupMap = this.getVoteGroupMap();
//
//        this.checkHaveGroup(groupList);
//        if (isHaveGroup) {
//            voteTypeGroupList.addAll(voteGroupMap.keySet());
//            for (String s : voteTypeGroupList) {
//                innerTitleList.add(voteGroupMap.get(s));
//            }
//            row = itemsList.size() + 4;
//            isHaveWBDS = this.checkWBDS(voterTypeList);
//            col = this.calculateColumn(voterTypeList, voteTypeGroupList, isHaveWBDS);
//        } else {
//            row = itemList.size() + 4;
//            col = voterTypeList.size() + 2;
//        }
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
//        String title = "{{title}}";
//        RowRenderData header0 = RowRenderData.build(new TextRenderData(title, cellStyle));
//        header0.setRowStyle(style);
//        TableTools.mergeCellsHorizonal(table, 0, 0, col - 1);
//        MiniTableRenderPolicy.Helper.renderRow(table, 0, header0);
//    }
//
//    // 设置第二、三行标题的样式
//    public void setTableHeader(XWPFTable table) {
//        if (isHaveGroup) {
//            // 含有分组
//            if ((voteTypeGroupList != null && voteTypeGroupList.size() > 0) && (innerTitleList != null && innerTitleList.size() > 0)) {
//                // 处理第二行，先判断第二行的表头内容
//                int length = this.voteTypeGroupList.size() + 2;
//                // 判断是否包含外部董事
//                if (isHaveWBDS) {
//                    length++;
//                }
//                String[] strHeader1 = new String[length];
//                strHeader1[0] = "评价内容";
//                strHeader1[length - 1] = "全体";
//                if (isHaveWBDS) {
//                    for (VoteRule rule : voterTypeList) {
//                        if ("".equals(rule.getGroup())) {
//                            strHeader1[length - 2] = rule.getFilterName() + "（" + rule.getRuleName() + "）";
//                        }
//                    }
//                }
//                for (int i = 0; i < voteTypeGroupList.size(); i++) {
//                    strHeader1[i + 1] = voteTypeGroupList.get(i);
//                }
//                // 构建第二行
//                Style cellStyle = this.getCellStyle();
//                RowRenderData header1 = this.build(strHeader1, cellStyle);
//                // 垂直合并 “全体”
//                TableTools.mergeCellsVertically(table, col - 1, 1, 2);
//                TableTools.mergeCellsVertically(table, col - 2, 1, 2);
//
//                // 是否含有外部董事
//                if (isHaveWBDS) {
//                    // 垂直合并 “外部董事”
//                    TableTools.mergeCellsVertically(table, col - 4, 1, 2);
//                }
//                // 垂直合并 “评价内容”
//                TableTools.mergeCellsVertically(table, 0, 1, 2);
//
//                // 水平合并 “评价内容”
//                TableTools.mergeCellsHorizonal(table, 1, 0, 1);
//                TableTools.mergeCellsHorizonal(table, 2, 0, 1);
//
//                // 水平合并 “分组”
//                int start = 1;
//                int end;
//                for (ArrayList<VoteRule> rules : innerTitleList) {
//                    end = (rules.size() + 1) * 2;
//                    TableTools.mergeCellsHorizonal(table, 1, start, start + end - 1);
//                    start++;
//                }
//
//                // 判断是否含有外部董事
//                if (isHaveWBDS) {
//                    // 水平合并 “外部董事”
//                    TableTools.mergeCellsHorizonal(table, 1, voteTypeGroupList.size() + 1, voteTypeGroupList.size() + 2);
//                    TableTools.mergeCellsHorizonal(table, 2, col - voteTypeGroupList.size() - 4, col - voteTypeGroupList.size() - 3);
//                    // 水平合并 “全体”
//                    TableTools.mergeCellsHorizonal(table, 1, voteTypeGroupList.size() + 2, voteTypeGroupList.size() + 3);
//                    TableTools.mergeCellsHorizonal(table, 2, col - voteTypeGroupList.size() - 3, col - voteTypeGroupList.size() - 2);
//                } else {
//                    TableTools.mergeCellsHorizonal(table, 1, voteTypeGroupList.size() + 1, voteTypeGroupList.size() + 2);
//                    TableTools.mergeCellsHorizonal(table, 2, col - voteTypeGroupList.size() - 1, col - voteTypeGroupList.size());
//                }
//                TableStyle tableStyle = this.getTableStyle();
//                header1.setRowStyle(tableStyle);
//                MiniTableRenderPolicy.Helper.renderRow(table, 1, header1);
//
//                // 处理分组内的票种和小计
//                int total = this.countRow(innerTitleList);
//                int lengthItem = (total + voteTypeGroupList.size()) + 1;
//                String[] strHeader2 = new String[lengthItem];
//                int index = 1;
//                for (ArrayList<VoteRule> rules : innerTitleList) {
//                    for (VoteRule rule : rules) {
//                        strHeader2[index++] = rule.getFilterName() + "（" + rule.getRuleName() + "）";
//                    }
//                    strHeader2[index++] = "小计";
//                }
//                RowRenderData header2 = this.build(strHeader2, cellStyle);
//                // 分别给以上几个 TextRenderData 设置合并单元格、第一个格子必须写空值否则无法写入到 table
//                for (int i = 0; i < strHeader2.length - 1; i++) {
//                    TableTools.mergeCellsHorizonal(table, 2, i + 1, i + 2);
//                }
//                header2.setRowStyle(tableStyle);
//                MiniTableRenderPolicy.Helper.renderRow(table, 2, header2);
//            } else {
//                int length = voterTypeList.size() + 2;
//                String[] strHeader1 = new String[length];
//                strHeader1[0] = "评价内容";
//                strHeader1[length - 1] = "全体";
//                for (int i = 0; i < voterTypeList.size(); i++) {
//                    strHeader1[i + 1] = voterTypeList.get(i).getFilterName();
//                }
//                // 构建第二行
//                Style cellStyle = this.getCellStyle();
//                RowRenderData header1 = this.build(strHeader1, cellStyle);
//                for (int i = 0; i < voterTypeList.size() + 2; i++) {
//                    int j = i + 1;
//                    TableTools.mergeCellsHorizonal(table, 1, i, j);
//                    TableTools.mergeCellsHorizonal(table, 2, i, j);
//                }
//
//                for (int i = voterTypeList.size() + 1; i >= 0; i--) {
//                    TableTools.mergeCellsVertically(table, i, 1, 2);
//                }
//                TableStyle tableStyle = this.getTableStyle();
//                header1.setRowStyle(tableStyle);
//                MiniTableRenderPolicy.Helper.renderRow(table, 1, header1);
//            }
//        } else {
//            int length = voterTypeList.size() + 2;
//            String[] strHeader1 = new String[length];
//            strHeader1[0] = "指标";
//            strHeader1[length - 1] = "全体";
//            for (int i = 0; i < voterTypeList.size(); i++) {
//                strHeader1[i + 1] = voterTypeList.get(i).getFilterName() + "（" + voterTypeList.get(i).getRuleName() + "）";
//            }
//            // 构建第二行
//            Style cellStyle = this.getCellStyle();
//            RowRenderData header1 = this.build(strHeader1, cellStyle);
//
//            // 垂直合并
//            for (int i = col - 1; i >= 0; i--) {
//                TableTools.mergeCellsVertically(table, i, 1, 2);
//            }
//            MiniTableRenderPolicy.Helper.renderRow(table, 1, header1);
//        }
//    }
//
//    // group 的设置、集中在第 1 列
//    public void setGroupAndItem(XWPFTable table) {
//        if (isHaveGroup) {
//            TableStyle tableStyle = this.getTableStyle();
//            Style cellStyle = this.getCellStyle();
//            String[] str = new String[col];
//            HashMap<String, ItemGroup> groupMap = new HashMap<>();
//            for (ItemGroup group : groupList) {
//                groupMap.put(group.getGroupID(), group);
//            }
//            for (int i = 0; i < itemsList.size(); i++) {
//                Item item = itemsList.get(i);
//                str[0] = groupMap.get(item.getGroupID()).getGroupName();
//                str[1] = item.getItemName();
//                RowRenderData rowData = this.build(str, cellStyle);
//                rowData.setRowStyle(tableStyle);
//                MiniTableRenderPolicy.Helper.renderRow(table, i + 3, rowData);
//            }
//            HashMap<String, Integer> map = new HashMap<>();
//            for (Item item : itemsList) {
//                if (!map.containsKey(item.getGroupID())) {
//                    map.put(item.getGroupID(), 0);
//                }
//                map.put(item.getGroupID(), map.get(item.getGroupID()) + 1);
//            }
//
//            int startRowIndex = 3;
//            for (ItemGroup group : groupList) {
//                int rowCount = map.get(group.getGroupID());
//                TableTools.mergeCellsVertically(table, 0, startRowIndex, startRowIndex + rowCount - 1);
//                startRowIndex += rowCount;
//            }
//        } else {
//            Style cellStyle = this.getCellStyle();
//            TableStyle tableStyle = this.getTableStyle();
//            for (int i = 0; i < itemList.size(); i++) {
//                RowRenderData itemData = RowRenderData.build(new TextRenderData(itemList.get(i).getItemName(), cellStyle));
//                itemData.setRowStyle(tableStyle);
//                MiniTableRenderPolicy.Helper.renderRow(table, i + 3, itemData);
//            }
//        }
//    }
//
//    // 设置 table 的 cell 中的 tag
//    public void setCellTag(XWPFTable table) {
//        if (isHaveGroup) {
//            String[] str = new String[col];
//            ArrayList<VoteRule> list = this.getAllRules();
//            HashMap<String, ItemGroup> groupMap = new HashMap<>();
//            for (ItemGroup group : groupList) {
//                groupMap.put(group.getGroupID(), group);
//            }
//            for (int i = 0; i < itemsList.size(); i++) {
//                Item item = itemsList.get(i);
//                for (int j = 2; j < str.length; j += 2) {
//                    if (j < str.length - 2) {
//                        if (list.get((j - 2) / 2).getVoterType().contains("rule")) {
//                            str[j] = "{{all100_" + item.getItemID() + "_" + list.get((j - 2) / 2).getVoterType().replaceAll(";", "") + "}}";
//                            str[j + 1] = "{{all100_" + groupMap.get(item.getGroupID()).getGroupID() + "_" + list.get((j - 2) / 2).getVoterType().replaceAll(";", "") + "}}";
//                        } else {
//                            str[j] = "{{avg_" + item.getItemID() + "_" + list.get((j - 2) / 2).getVoterType().replaceAll(";", "") + "}}";
//                            str[j + 1] = "{{avg_" + groupMap.get(item.getGroupID()).getGroupID() + "_" + list.get((j - 2) / 2).getVoterType().replaceAll(";", "") + "}}";
//                        }
//                    } else {
//                        if (j == str.length - 2) {
//                            str[j] = "{{avg_" + item.getItemID() + "}}";
//                            str[j + 1] = "{{avg_" + groupMap.get(item.getGroupID()).getGroupID() + "}}";
//                        }
//                    }
//                }
//                Style cellStyle = this.getDataCellStyle();
//                RowRenderData rowData = this.build(str, cellStyle);
//                TableStyle tableStyle = this.getTableStyle();
//                rowData.setRowStyle(tableStyle);
//                MiniTableRenderPolicy.Helper.renderRow(table, i + 3, rowData);
//            }
//
//            HashMap<String, Integer> map = new HashMap<>();
//            for (Item item : itemsList) {
//                if (!map.containsKey(item.getGroupID())) {
//                    map.put(item.getGroupID(), 0);
//                }
//                map.put(item.getGroupID(), map.get(item.getGroupID()) + 1);
//            }
//
//            int startRowIndex = 3;
//            for (ItemGroup group : groupList) {
//                int rowCount = map.get(group.getGroupID());
//                int column = 3;
//                for (int i = 0; i < col / 2 - 1; i++) {
//                    TableTools.mergeCellsVertically(table, column, startRowIndex, startRowIndex + rowCount - 1);
//                    column += 2;
//                }
//                startRowIndex += rowCount;
//            }
//        } else {
//            TableStyle tableStyle = this.getTableStyle();
//            Style cellStyle = this.getCellStyle();
//            String[] str = new String[col];
//            for (int i = 0; i < itemList.size(); i++) {
//                Item item = itemList.get(i);
//                for (int j = 1; j < str.length; j++) {
//                    str[j] = "{{avg#" + item.getItemID() + "@" + voterTypeList.get(i - 1) + "}}";
//                    if (j == str.length - 1) {
//                        str[j] = "{{avg#" + item.getItemID() + "}}";
//                    }
//                }
//                RowRenderData rowData = this.build(str, cellStyle);
//                rowData.setRowStyle(tableStyle);
//                MiniTableRenderPolicy.Helper.renderRow(table, i + 3, rowData);
//            }
//        }
//    }
//
//    // 设置最后一行
//    public void setLastRow(XWPFTable table) {
//        if (isHaveGroup) {
//            // 最后一行的加权总得分
//            String[] str = new String[col / 2];
//            str[0] = "加权汇总得分";
//            ArrayList<VoteRule> list = this.getAllRules();
//            for (int i = 0; i < list.size(); i++) {
//                if (list.get(i).getVoterType().contains("rule")) {
//                    str[i + 1] = "{{all100_" + list.get(i).getVoterType().replaceAll(";", "") + "}}";
//                } else {
//                    str[i + 1] = "{{avg_" + list.get(i).getVoterType().replaceAll(";", "") + "}}";
//                }
//            }
//            str[str.length - 1] = "{{avg}}";
//            Style cellStyle = this.getCellStyle();
//            RowRenderData total = this.build(str, cellStyle);
//            for (int i = 0; i < col / 2; i++) {
//                TableTools.mergeCellsHorizonal(table, row - 1, i, i + 1);
//            }
//            TableStyle tableStyle = this.getTableStyle();
//            total.setRowStyle(tableStyle);
//            MiniTableRenderPolicy.Helper.renderRow(table, row - 1, total);
//        } else {
//            // 最后一行的加权总得分
//            String[] str = new String[col];
//            str[0] = "加权汇总得分";
//            for (int i = 0; i < voterTypeList.size(); i++) {
//                str[i + 1] = "{{avg@" + voterTypeList.get(i).getVoterType() + "}}";
//            }
//            str[str.length - 1] = "{{avg}}";
//            Style cellStyle = this.getCellStyle();
//            RowRenderData total = this.build(str, cellStyle);
//            TableStyle tableStyle = this.getTableStyle();
//            total.setRowStyle(tableStyle);
//            MiniTableRenderPolicy.Helper.renderRow(table, row - 1, total);
//        }
//    }
//
//    // 判断是否包含一级指标
//    public void checkHaveGroup(ArrayList<ItemGroup> groupList) {
//        for (ItemGroup itemGroup : groupList) {
//            // 默认包含一级分组
//            isHaveGroup = true;
//            // 如果 groupList 中有 groupName 为默认分组的则为没有一级指标
//            if ("默认分组".equals(itemGroup.getGroupName())) {
//                isHaveGroup = false;
//                break;
//            }
//        }
//    }
//
//    // 判断否含有外部董事
//    public boolean checkWBDS(List<VoteRule> voterTypeList) {
//        isHaveWBDS = false;
//        Set<String> set = new HashSet<>();
//        for (VoteRule rule : voterTypeList) {
//            set.add(rule.getGroup());
//        }
//        if (set.size() > 1) {
//            if (set.contains("")) {
//                isHaveWBDS = true;
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
//    // 组名和分组对应
//    public LinkedHashMap<String, ArrayList<VoteRule>> getVoteGroupMap() throws DAOException {
//        LinkedHashMap<String, ArrayList<VoteRule>> voteGroupMap = new LinkedHashMap<>();
//        ArrayList<VoteRule> ruleList = VoteRuleDAO.getVoteRuleListBySystem(system.getSystemID());
//        for (VoteRule rule : ruleList) {
//            if (rule.getGroup().length() > 0) {
//                voteGroupMap.put(rule.getGroup(), VoteRuleDAO.getVoteRuleListBySQL("where groupname = '" + rule.getGroup() + "' and systemid='" + system.getSystemID() + "'"));
//            }
//        }
//        return voteGroupMap;
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
//
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
//}*/
