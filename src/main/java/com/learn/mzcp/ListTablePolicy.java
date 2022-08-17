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
//import com.first.leaderfieldconfig.LeaderFieldConfig;
//import com.first.leaderfieldconfig.LeaderFieldConfigDAO;
//import com.first.voteelement.VoteElement;
//import com.first.voterule.VoteRuleSystem;
//import org.apache.poi.xwpf.usermodel.*;
//import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJc;
//
//import java.util.ArrayList;
//import java.util.List;
//
///**
// * @author ：Kristen
// * @date ：2022/7/12
// * @description :
// */
//public class ListTablePolicy extends AbstractRenderPolicy<Object> {
//
//    private VoteElement element;
//    private VoteRuleSystem system;
//    private int dataSize;
//
//    // groupList
//    public ArrayList<ItemGroup> groupList;
//    // itemsList（有一级指标的时候）
//    public ArrayList<Item> itemsList;
//    // 根据 group 分组的 itemsGroupList
//    public ArrayList<ArrayList<Item>> itemsGroupList = new ArrayList<>();
//    // itemList（没有一级指标的时候）
//    public ArrayList<Item> itemList;
//    // 读取配置 list
//    public ArrayList<LeaderFieldConfig> configList;
//
//    // 计算行和列
//    int col;
//    int row;
//    // 判断是否包含一级指标
//    boolean isHaveGroup;
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
//        // 左侧第一列 group
//        groupList = ItemGroupDAO.getGroupListByTable(element.getTableID());
//        // 左侧第二列 items（有一级指标的时候）
//        itemsList = ItemDAO.getItemListByTable(element.getTableID());
//        // 左侧第一列 item（没有一级指标的时候）
//        itemList = ItemDAO.getItemListByTable(element.getTableID());
//        itemsGroupList = this.getItemsGroupList(groupList, itemsList);
//
//        // 获取配置 list
//        String backSQL = "where menuIndex = '2' and usedelements is not null";
//        configList = LeaderFieldConfigDAO.getLeaderFieldConfigListBySQL(backSQL);
//
//        this.checkHaveGroup(groupList);
//        if (isHaveGroup) {
//            col = this.countCol(itemsGroupList) + configList.size() + 3;
//        } else {
//            col = itemList.size() + configList.size() + 3;
//        }
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
//        if (isHaveGroup) {
//            // 构建第二行的数组、并垂直合并与水平合并
//            String[] strHeader1 = new String[groupList.size() + configList.size() + 3];
//            strHeader1[0] = "序号";
//            for (int i = 0; i < configList.size(); i++) {
//                strHeader1[i + 1] = configList.get(i).getFieldName();
//            }
//            strHeader1[configList.size() + 1] = "汇总得分";
//            strHeader1[configList.size() + 2] = "排名";
//
//            int start = configList.size() + 3;
//            int end;
//            // 水平合并 group 的名字
//            for (ArrayList<Item> items : itemsGroupList) {
//                end = (items.size() + 1);
//                TableTools.mergeCellsHorizonal(table, 1, start, start + end - 1);
//                start++;
//            }
//
//            for (int i = 0; i < groupList.size(); i++) {
//                strHeader1[configList.size() + i + 3] = groupList.get(i).getGroupName();
//                TableTools.mergeCellsVertically(table, i, 1, 2);
//            }
//
//            // 构建第三行的数组
//            String[] strHeader2 = new String[col];
//            int index = configList.size() + 3;
//            for (ArrayList<Item> items : itemsGroupList) {
//                strHeader2[index] = "小计";
//                for (Item item : items) {
//                    strHeader2[++index] = item.getItemName();
//                }
//                index++;
//            }
//
//            // 构建第二、三行
//            Style style = this.getDataCellStyle();
//            RowRenderData header1 = this.build(strHeader1, style);
//            RowRenderData header2 = this.build(strHeader2, style);
//            TableStyle tableStyle = this.getTableStyle();
//            header1.setRowStyle(tableStyle);
//            header2.setRowStyle(tableStyle);
//            MiniTableRenderPolicy.Helper.renderRow(table, 1, header1);
//            MiniTableRenderPolicy.Helper.renderRow(table, 2, header2);
//        } else {
//            String[] strHeader1 = new String[itemList.size() + configList.size() + 3];
//            strHeader1[0] = "序号";
//            for (int i = 0; i < configList.size(); i++) {
//                strHeader1[i + 1] = configList.get(i).getFieldName();
//            }
//            strHeader1[configList.size() + 1] = "汇总得分";
//            strHeader1[configList.size() + 2] = "排名";
//            for (int i = 0; i < itemList.size(); i++) {
//                strHeader1[i + 5] = itemList.get(i).getItemName();
//            }
//            Style style = this.getDataCellStyle();
//            RowRenderData header1 = this.build(strHeader1, style);
//            TableStyle tableStyle = this.getTableStyle();
//            header1.setRowStyle(tableStyle);
//            MiniTableRenderPolicy.Helper.renderRow(table, 1, header1);
//        }
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
//            str[configList.size() + 1] = "{{avg_" + k + "}}";
//            str[configList.size() + 2] = "{{sort_avg_" + k + "}}";
//
//            int index = configList.size() + 3;
//            if (groupList.size() > 0) {
//                for (int i = 0; i < groupList.size(); i++) {
//                    str[index] = "{{avg_group0" + (i + 1) + "_" + k + "}}";
//                    for (int j = 0; j < itemsGroupList.get(i).size(); j++) {
//                        if (index - 5 - i < 9) {
//                            str[++index] = "{{avg_leader0" + (index - 5 - i) + "_" + k + "}}";
//                        } else {
//                            str[++index] = "{{avg_leader" + (index - 5 - i) + "_" + k + "}}";
//                        }
//                    }
//                    index++;
//                }
//            } else {
//                for (int i = 0; i < itemList.size(); i++) {
//                    if (i < 9) {
//                        str[i + index] = "{{avg_leader0" + (i + 1) + "_" + k + "}}";
//                    } else {
//                        str[i + index] = "{{avg_leader" + (i + 1) + "_" + k + "}}";
//                    }
//                }
//            }
//            Style style = this.getDataCellStyle();
//            RowRenderData row = this.build(str, style);
//            TableStyle tableStyle = this.getTableStyle();
//            row.setRowStyle(tableStyle);
//            MiniTableRenderPolicy.Helper.renderRow(table, k + 3, row);
//        }
//    }
//
//    public ArrayList<ArrayList<Item>> getItemsGroupList(ArrayList<ItemGroup> groupList, ArrayList<Item> itemsList) {
//        for (ItemGroup itemGroup : groupList) {
//            ArrayList<Item> list = new ArrayList<>();
//            for (Item item : itemsList) {
//                if (itemGroup.getGroupID().equals(item.getGroupID())) {
//                    list.add(item);
//                }
//            }
//            itemsGroupList.add(list);
//        }
//        return itemsGroupList;
//    }
//
//    // 计算所有分组的项的个数
//    public int countCol(ArrayList<ArrayList<Item>> itemsGroupList) {
//        int count = 0;
//        int index = 0;
//        for (ArrayList<Item> items : itemsGroupList) {
//            index++;
//            for (int j = 0; j < items.size(); j++) {
//                count++;
//            }
//        }
//        return count + index;
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
