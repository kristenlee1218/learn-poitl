package com.learn.crossTableTemplate;

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
import java.util.HashSet;
import java.util.List;
import java.util.Set;

/**
 * @author ：Kristen
 * @date ：2022/7/14
 * @description : 此策略中，可以生成班子汇总（包括无一级指标的）和人员（单个）汇总的交叉表
 */
public class CrossTablePolicy extends AbstractRenderPolicy<Object> {

    // 第一种测试情况（信息中心_2020年第1批综合测评统计报表(1655863588387)）(14)
//    public static String[] group = new String[]{"政治思想建设", "企业发展质量", "党建工作质量", "作风建设成效"};
//    public static String[][] items = new String[][]{{"政治忠诚", "政治担当", "社会责任"}, {"改革创新", "经营效益", "管理效能", "风险管控"}, {"选人用人", "基层党建", "党风廉政"}, {"团结协作", "联系群众"}};
//    public static String[] item = new String[]{};
//    public static String[] voteType = new String[]{"领导班子成员（A1、A2、A3票）", "中层测评（B票）", "职工代表（C票）", "外部董事（A4票）"};
//    public static String[] voteTypeGroup = new String[]{"内部测评"};
//    public static String[][] innerEvaluate = new String[][]{{"领导班子成员（A1、A2、A3票）", "中层测评（B票）", "职工代表（C票）"}};
//    public static String[] value = new String[]{"1.00", "2.00", "3.00", "4.00", "5.00", "6.00", "7.00", "8.00", "9.00", "10.00", "11.00", "12.00", "13.00", "14.00", "15.00", "16.00", "17.00", "18.00", "19.00", "20", "21.00", "22.00", "23.00", "24.00", "25.00", "26.00", "27.00", "28.00", "29.00", "30.00", "31.00", "32.00", "33", "34.00", "35.00", "36.00", "37.00", "38.00", "39.00", "40.00", "41.00", "42.00", "43.00", "44.00", "45.00", "46", "47.00", "48.00", "49.00", "50.00", "51.00", "52.00", "53.00", "54.00", "55.00", "56.00", "57.00", "58.00", "59", "60.00", "61.00", "62.00", "63.00", "64.00", "65.00", "66.00", "67.00", "68.00", "69.00", "70.00", "71.00", "72", "73.00", "74.00", "75.00", "76.00", "77.00", "78.00", "79.00", "80.00", "81.00", "82.00", "83.00", "84.00", "85", "86.00", "87.00", "88.00", "89.00", "90.00", "91.00", "92.00", "93.00", "94.00", "95.00", "96.00", "97.00", "98", "99.00", "100.00", "101.00", "102.00"};

    //    第二种测试情况（上海大区公司_2021年第1批综合考核评价统计表）(20)
//    public static String[] group = new String[]{"政治思想建设", "企业发展质量", "党建工作质量", "作风建设成效"};
//    public static String[][] items = new String[][]{{"政治忠诚", "政治担当", "社会责任"}, {"改革创新", "经营效益"}, {"选人用人", "基层党建", "党风廉政"}, {"团结协作", "联系群众"}};
//    public static String[] item = new String[]{};
//    public static String[] voteType = new String[]{"董事长(A1)", "总经理(A2)", "其他领导(A3)", "本单位领导班子成员(B1、B2)", "中层经理人(C)", "员工(D)"};
//    public static String[] voteTypeGroup = new String[]{"大悦城控股领导", "内部测评"};
//    public static String[][] innerEvaluate = new String[][]{{"董事长(A1)", "总经理(A2)", "其他领导(A3)"}, {"本单位领导班子成员(B1、B2)", "中层经理人(C)", "员工(D)"}};
//    public static String[] value = new String[]{"1.00", "2.00", "3.00", "4.00", "5.00", "6.00", "7.00", "8.00", "9.00", "10.00", "11.00", "12.00", "13.00", "14.00", "15.00", "16.00", "17.00", "18.00", "19.00", "20.00", "21.00", "22.00", "23.00", "24.00", "25.00", "26.00", "27.00", "28.00", "29.00", "30.00", "31.00", "32.00", "33.00", "34.00", "35.00", "36.00", "37.00", "38.00", "39.00", "40.00", "41.00", "42.00", "43.00", "44.00", "45.00", "46.00", "47.00", "48.00", "49.00", "50.00", "51.00", "52.00", "53.00", "54.00", "55.00", "56.00", "57.00", "58.00", "59.00", "60.00", "61.00", "62.00", "63.00", "64.00", "65.00", "66.00", "67.00", "68.00", "69.00", "70.00", "71.00", "72.00", "73.00", "74.00", "75.00", "76.00", "77.00", "78.00", "79.00", "80.00", "81.00", "82.00", "83.00", "84.00", "85.00", "86.00", "87.00", "88.00", "89.00", "90.00", "91.00", "92.00", "93.00", "94.00", "95.00", "96.00", "97.00", "98.00", "99.00", "100.00", "101.00", "102.00", "103.00", "104.00", "105.00", "106.00", "107.00", "108.00", "109.00", "110.00", "111.00", "112.00", "113.00", "114.00", "115.00", "116.00", "117.00", "118.00", "119.00", "120.00", "121.00", "122.00", "123.00", "124.00", "125.00", "126.00", "127.00", "128.00", "129.00", "130.00", "131.00", "132.00", "133.00", "134.00", "135.00"};

    //    第三种测试情况（上海_2021年第1批综合考核评价统计报表(1646623484403)）（10）
//    public static String[] group = new String[]{"政治思想建设", "企业发展质量", "党建工作质量", "作风建设成效"};
//    public static String[][] items = new String[][]{{"政治忠诚", "政治担当", "社会责任"}, {"改革创新", "经营效益"}, {"选人用人", "基层党建", "党风廉政"}, {"团结协作", "联系群众"}};
//    public static String[] item = new String[]{};
//    public static String[] voteType = new String[]{"A票", "B票", "C票"};
//    public static String[] voteTypeGroup = new String[]{};
//    public static String[][] innerEvaluate = new String[][]{{}};
//    public static String[] value = new String[]{"1.00", "2.00", "3.00", "4.00", "5.00", "6.00", "7.00", "8.00", "9.00", "10.00", "11.00", "12.00", "13.00", "14.00", "15.00", "16.00", "17.00", "18.00", "19.00", "20.00", "21.00", "22.00", "23.00", "24.00", "25.00", "26.00", "27.00", "28.00", "29.00", "30.00", "31.00", "32.00", "33.00", "34.00", "35.00", "36.00", "37.00", "38.00", "39.00", "40.00", "41.00", "42.00", "43.00", "44.00", "45.00", "46.00", "47.00", "48.00", "49.00", "50.00", "51.00", "52.00", "53.00", "54.00", "55.00", "56.00", "57.00", "58.00", "59.00", "60.00"};

    //    第四种测试情况（河北建投二级单位_2022年第1批综合考核评价统计报表(1653979325662)）
//    public static String[] group = new String[]{};
//    public static String[][] items = new String[][]{{}};
//    public static String[] item = new String[]{"政治方向", "社会责任", "企业党建", "科学管理", "发扬民主", "整体合力", "诚信合力", "联系群众", "廉洁自律", "工作部署"};
//    public static String[] voteType = new String[]{"A票", "B票"};
//    public static String[] voteTypeGroup = new String[]{};
//    public static String[][] innerEvaluate = new String[][]{{}};
//    public static String[] value = new String[]{"1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31", "32", "33"};

    // 第五种测试情况（领导人员）
    public static String[] group = new String[]{"对党忠诚", "勇于创新", "治企有方", "兴企有为", "清正廉洁"};
    public static String[][] items = new String[][]{{"政治品质", "政治本领"}, {"创新精神", "创新成果"}, {"经营管理能力", "抓党建强党建能力"}, {"担当作为", "履职绩效"}, {"一岗双责", "廉洁从业"}};
    public static String[] item = new String[]{};
    public static String[] voteType = new String[]{"领导班子成员\n（A1、A2、A3票）", "中层测评\n（B票）", "职工代表\n（C票）", "外部董事\n（A4票）"};
    public static String[] voteTypeGroup = new String[]{"内部测评"};
    public static String[][] innerEvaluate = new String[][]{{"领导班子成员\n（A1、A2、A3票）", "中层测评\n（B票）", "职工代表\n（C票）"}};
    public static String[] value = new String[]{"1.00", "2.00", "3.00", "4.00", "5.00", "6.00", "7.00", "8.00", "9.00", "10.00", "11.00", "12.00", "13.00", "14.00", "15.00", "16.00", "17.00", "18.00", "19.00", "20.00", "21.00", "22.00", "23.00", "24.00", "25.00", "26.00", "27.00", "28.00", "29.00", "30.00", "31.00", "32.00", "33", "34.00", "35.00", "36.00", "37.00", "38.00", "39.00", "40.00", "41.00", "42.00", "43.00", "44.00", "45.00", "46", "47.00", "48.00", "49.00", "50.00", "51.00", "52.00", "53.00", "54.00", "55.00", "56.00", "57.00", "58.00", "59", "60.00", "61.00", "62.00", "63.00", "64.00", "65.00", "66.00", "67.00", "68.00", "69.00", "70.00", "71.00", "72", "73.00", "74.00", "75.00", "76.00", "77.00", "78.00", "79.00", "80.00", "81.00", "82.00", "83.00", "84.00", "85", "86.00", "87.00", "88.00", "89.00", "90.00", "91.00", "92.00", "93.00", "94.00", "95.00", "96.00", "97.00", "98.00", "99.00", "100.00", "101.00", "102.00"};

    // 计算行
    int row;

    // 计算列、先判断有无外部董事
    boolean isHaveWBDS;
    int col;

    // 标题+表头所占的固定行数
    int rowBase = 3;
    // group 和 item 所占的固定列数
    int colBase = 2;

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

        // 计算行
        row = this.getRow(group);

        // 计算列、先判断有无外部董事
        isHaveWBDS = this.checkWBDS(voteType);
        col = this.getCol(group);

        // 当前位置插入表格
        XWPFTable table = bodyContainer.insertNewTable(run, row, col);

        this.setTableStyle(table);
        this.setTableTitle(table);
        this.setTableHeader(table);
        if (group.length > 0) {
            this.setGroup(table);
        }
        this.setItem(table);
        this.setLastRow(table);
        this.setTableData(table);
    }

    // 整个 table 的样式在此设置
    public void setTableStyle(XWPFTable table) {
        // 设置 A4 幅面的平铺类型和列数
        if (col <= 15) {
            TableTools.widthTable(table, MiniTableRenderData.WIDTH_A4_NARROW_FULL, col);
        } else {
            TableTools.widthTable(table, MiniTableRenderData.WIDTH_A4_NARROW_FULL, col);
        }
        TableTools.borderTable(table, 1);
        for (XWPFTableRow tableRow : table.getRows()) {
            for (int i = 0; i < tableRow.getTableCells().size(); i++) {
                tableRow.getCell(i).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
                tableRow.getCell(i).setWidth("500");
            }
        }
        // 设置一个 table 的边框值
        table.setBottomBorder(XWPFTable.XWPFBorderType.SINGLE, 20, 0, "000000");
        table.setTopBorder(XWPFTable.XWPFBorderType.SINGLE, 20, 0, "000000");
        table.setLeftBorder(XWPFTable.XWPFBorderType.SINGLE, 20, 0, "000000");
        table.setRightBorder(XWPFTable.XWPFBorderType.SINGLE, 20, 0, "000000");
        table.setCellMargins(5, 10, 5, 10);
        table.setTableAlignment(TableRowAlign.CENTER);
    }

    // 设置第一行标题的样式
    public void setTableTitle(XWPFTable table) {
        Style cellStyle = new Style();
        cellStyle.setFontSize(12);
        cellStyle.setFontFamily("黑体");
        cellStyle.setColor("000000");
        TableStyle tableStyle = new TableStyle();
        tableStyle.setAlign(STJc.CENTER);
        tableStyle.setBackgroundColor("F5F5F5");
        String title = "{{title}}";
        RowRenderData header0 = RowRenderData.build(new TextRenderData(title, cellStyle));
        header0.setRowStyle(tableStyle);
        TableTools.mergeCellsHorizonal(table, 0, 0, col - 1);
        MiniTableRenderPolicy.Helper.renderRow(table, 0, header0);
    }

    // 设置第二、三行标题的样式
    public void setTableHeader(XWPFTable table) {
        // 有一级指标并含有分组
        if (group.length > 0) {
            if ((voteTypeGroup != null && voteTypeGroup.length > 0) && (innerEvaluate != null && innerEvaluate.length > 0)) {
                // 处理第二行，先判断第二行的表头内容
                int length = voteTypeGroup.length + colBase;
                // 判断是否包含外部董事
                if (isHaveWBDS) {
                    length++;
                }
                String[] strHeader1 = new String[length];
                strHeader1[0] = "评价内容";
                strHeader1[length - 1] = "全体";
                if (isHaveWBDS) {
                    strHeader1[length - 2] = "外部董事（A4票）";
                }
                System.arraycopy(voteTypeGroup, 0, strHeader1, 1, voteTypeGroup.length);

                // 构建第二行
                Style cellStyle = this.getCellStyle();
                RowRenderData header1 = this.build(strHeader1, cellStyle);
                // 垂直合并"全体"
                TableTools.mergeCellsVertically(table, col - 1, 1, 2);
                TableTools.mergeCellsVertically(table, col - 2, 1, 2);

                // 是否含有外部董事
                if (isHaveWBDS) {
                    // 垂直合并“外部董事”
                    TableTools.mergeCellsVertically(table, col - 4, 1, 2);
                }
                // 垂直合并“评价内容”
                TableTools.mergeCellsVertically(table, 0, 1, 2);

                // 水平合并“评价内容”
                TableTools.mergeCellsHorizonal(table, 1, 0, 1);
                TableTools.mergeCellsHorizonal(table, 2, 0, 1);

                // 水平合并“分组”
                int start = 1;
                int end;
                for (String[] value : innerEvaluate) {
                    end = (value.length + 1) * 2;
                    TableTools.mergeCellsHorizonal(table, 1, start, start + end - 1);
                    start++;
                }

                // 判断是否含有外部董事
                if (isHaveWBDS) {
                    // 水平合并“外部董事”
                    TableTools.mergeCellsHorizonal(table, 1, voteTypeGroup.length + 1, voteTypeGroup.length + 2);
                    TableTools.mergeCellsHorizonal(table, 2, col - voteTypeGroup.length - 4, col - voteTypeGroup.length - 3);
                    // 水平合并“全体”
                    TableTools.mergeCellsHorizonal(table, 1, voteTypeGroup.length + 2, voteTypeGroup.length + 3);
                    TableTools.mergeCellsHorizonal(table, 2, col - voteTypeGroup.length - 3, col - voteTypeGroup.length - 2);
                } else {
                    TableTools.mergeCellsHorizonal(table, 1, voteTypeGroup.length + 1, voteTypeGroup.length + 2);
                    TableTools.mergeCellsHorizonal(table, 2, col - voteTypeGroup.length - 1, col - voteTypeGroup.length);
                }
                TableStyle tableStyle = this.getTableStyle();
                header1.setRowStyle(tableStyle);
                MiniTableRenderPolicy.Helper.renderRow(table, 1, header1);

                // 处理分组内的票种和小计
                int total = countRow(innerEvaluate);
                int lengthItem = (total + voteTypeGroup.length) + 1;
                String[] strHeader2 = new String[lengthItem];
                int index = 1;
                for (String[] strings : innerEvaluate) {
                    for (String str : strings) {
                        strHeader2[index++] = str;
                    }
                    strHeader2[index++] = "小计";
                }
                RowRenderData header2 = this.build(strHeader2, cellStyle);
                // 分别给以上几个 TextRenderData 设置合并单元格、第一个格子必须写空值否则无法写入到 table
                for (int i = 0; i < strHeader2.length - 1; i++) {
                    TableTools.mergeCellsHorizonal(table, 2, i + 1, i + 2);
                }
                header2.setRowStyle(tableStyle);
                MiniTableRenderPolicy.Helper.renderRow(table, 2, header2);
            } else {
                Style cellStyle = this.getCellStyle();
                String[] strHeader1 = new String[voteType.length + colBase];
                strHeader1[0] = "评价内容";
                strHeader1[voteType.length + colBase - 1] = "全体";
                System.arraycopy(voteType, 0, strHeader1, 1, voteType.length);
                // 构建第二行
                RowRenderData header1 = this.build(strHeader1, cellStyle);
                for (int i = 0; i < voteType.length + 2; i++) {
                    int j = i + 1;
                    TableTools.mergeCellsHorizonal(table, 1, i, j);
                    TableTools.mergeCellsHorizonal(table, 2, i, j);
                }
                for (int i = voteType.length + 1; i >= 0; i--) {
                    TableTools.mergeCellsVertically(table, i, 1, 2);
                }
                TableStyle tableStyle = this.getTableStyle();
                header1.setRowStyle(tableStyle);
                MiniTableRenderPolicy.Helper.renderRow(table, 1, header1);
            }
        } else {
            String[] strHeader1 = new String[voteType.length + colBase];
            strHeader1[0] = "指标";
            strHeader1[voteType.length + colBase - 1] = "全体";
            System.arraycopy(voteType, 0, strHeader1, 1, voteType.length);

            // 构建第二行
            Style cellStyle = this.getCellStyle();
            RowRenderData header1 = this.build(strHeader1, cellStyle);

            // 垂直合并
            for (int i = col - 1; i >= 0; i--) {
                TableTools.mergeCellsVertically(table, i, 1, 2);
            }
            TableStyle tableStyle = this.getTableStyle();
            header1.setRowStyle(tableStyle);
            MiniTableRenderPolicy.Helper.renderRow(table, 1, header1);
        }
    }

    // group 的设置、集中在第1列
    public void setGroup(XWPFTable table) {
        int start = rowBase;
        int end;
        Style cellStyle = this.getCellStyle();
        TableStyle tableStyle = this.getTableStyle();
        for (int i = 0; i < items.length; i++) {
            end = start + items[i].length;
            TableTools.mergeCellsVertically(table, 0, start, end - 1);
            RowRenderData groupData = RowRenderData.build(new TextRenderData(group[i], cellStyle));
            groupData.setRowStyle(tableStyle);
            MiniTableRenderPolicy.Helper.renderRow(table, start, groupData);
            start = end;
        }
    }

    // item 的设置、集中在第2列
    public void setItem(XWPFTable table) {
        // 有一级指标并含有分组
        if (group.length > 0) {
            // item 的设置、集中在第 2 列
            // 构建 item 列
            int index = rowBase;
            Style cellStyle = this.getCellStyle();
            Style cell8Style = this.getCell8Style();
            Style cell6Style = this.getCell6Style();
            TableStyle tableStyle = this.getTableStyle();
            for (String[] str : items) {
                for (String s : str) {
                    RowRenderData itemData;
                    if (s.length() > 4 && s.length() < 8) {
                        itemData = RowRenderData.build(new TextRenderData("", cellStyle), new TextRenderData(s, cell8Style));
                    } else if (s.length() >= 8) {
                        itemData = RowRenderData.build(new TextRenderData("", cellStyle), new TextRenderData(s, cell6Style));
                    } else {
                        itemData = RowRenderData.build(new TextRenderData("", cellStyle), new TextRenderData(s, cellStyle));
                    }
//                    itemData = RowRenderData.build(new TextRenderData("", cellStyle), new TextRenderData(s, cellStyle));
                    itemData.setRowStyle(tableStyle);
                    MiniTableRenderPolicy.Helper.renderRow(table, index++, itemData);
                }
            }

            // 合并 item 列
            int column = rowBase;
            for (int i = 0; i < col / 2 - 1; i++) {
                int start = rowBase;
                int end;
                for (String[] str : items) {
                    end = start + str.length;
                    // 合并 item 均分的单元格
                    TableTools.mergeCellsVertically(table, column, start, end - 1);
                    start = end;
                }
                column += 2;
            }
        } else {
            TableStyle tableStyle = this.getTableStyle();
            Style cellStyle = this.getCellStyle();
            for (int i = 0; i < item.length; i++) {
                RowRenderData itemData = RowRenderData.build(new TextRenderData(item[i], cellStyle));
                itemData.setRowStyle(tableStyle);
                MiniTableRenderPolicy.Helper.renderRow(table, i + rowBase, itemData);
            }
        }
    }

    // 设置最后一行
    public void setLastRow(XWPFTable table) {
        // 最后一行的加权总得分
        Style cellStyle = this.getCellStyle();
        Style cell9Style = this.getCell9Style();
        RowRenderData total = RowRenderData.build(new TextRenderData("加权汇总得分", cell9Style));
        TableStyle tableStyle = this.getTableStyle();
        total.setRowStyle(tableStyle);
        if (group.length > 0) {
            for (int i = 0; i < col / 2; i++) {
                TableTools.mergeCellsHorizonal(table, row - 1, i, i + 1);
            }
        }
        MiniTableRenderPolicy.Helper.renderRow(table, row - 1, total);
    }

    // 将 table 所需的数据值填充到对应的格子中
    public void setTableData(XWPFTable table) {
        if (group.length > 0) {
            // 设置计算分数的单元格值（除最后一行）
            int index = 0;
            int start = rowBase;
            Set<Integer> set = this.calculateGroupStartRow(items, start);
            Style cellStyle = this.getDataCellStyle();
            TableStyle tableStyle = this.getTableStyle();
            for (int i = start; i < row - 1; i++) {
                String[] str = new String[col];
                if (set.contains(i)) {
                    for (int j = 2; j < table.getRow(i).getTableCells().size(); j++) {
                        str[j] = value[index];
                        index++;
                    }
                    RowRenderData lineValue = this.build(str, cellStyle);
                    lineValue.setRowStyle(tableStyle);
                    MiniTableRenderPolicy.Helper.renderRow(table, i, lineValue);
                } else {
                    for (int j = 2; j < table.getRow(i).getTableCells().size(); j += 2) {
                        str[j] = value[index];
                        index++;
                    }
                    RowRenderData lineValue = this.build(str, cellStyle);
                    lineValue.setRowStyle(tableStyle);
                    MiniTableRenderPolicy.Helper.renderRow(table, i, lineValue);
                }
            }
            // 最后一行的值
            String[] str = new String[col / 2];
            for (int i = 1; i < col / 2; i++) {
                str[str.length - i] = value[value.length - i];
            }
            RowRenderData lineValue = this.build(str, cellStyle);
            lineValue.setRowStyle(tableStyle);
            MiniTableRenderPolicy.Helper.renderRow(table, row - 1, lineValue);
        } else {
            Style cellStyle = this.getCellStyle();
            TableStyle tableStyle = this.getTableStyle();
            int start = rowBase;
            int index = 0;
            for (int i = start; i < row; i++) {
                String[] str = new String[col];
                for (int j = 1; j < table.getRow(i).getTableCells().size(); j++) {
                    str[j] = value[index];
                    index++;
                }
                RowRenderData lineValue = this.build(str, cellStyle);
                lineValue.setRowStyle(tableStyle);
                MiniTableRenderPolicy.Helper.renderRow(table, i, lineValue);
            }
        }
    }

    // 计算行（区分有没有一级指标）
    public int getRow(String[] group) {
        // 有一级指标和二级指标
        if (group.length > 0) {
            row = this.countRow(items) + 4;
        } else {
            row = item.length + 4;
        }
        return row;
    }

    // 计算列（区分有没有一级指标）
    public int getCol(String[] group) {
        // 有一级指标和二级指标
        if (group.length > 0) {
            isHaveWBDS = this.checkWBDS(voteType);
            col = this.calculateColumn(voteType, voteTypeGroup, isHaveWBDS);
        } else {
            col = voteType.length + 2;
        }
        return col;
    }

    // 判断否含有外部董事
    public boolean checkWBDS(String[] voteType) {
        for (String str : voteType) {
            if (str.contains("外部董事")) {
                isHaveWBDS = true;
                break;
            }
        }
        return isHaveWBDS;
    }

    // 计算列数
    public int calculateColumn(String[] voteType, String[] voteTypeGroup, boolean isHaveWBDS) {
        int column;
        int littleCount = 0;
        // 如果分组大于 1，则无论有无外部董事，每个组都有一个小计
        if (voteTypeGroup.length > 1) {
            littleCount = voteTypeGroup.length * 2;
        } else {
            // 只有一个分组的情况下，如果有外部董事则有一个小计，如果没有外部董事则没有小计
            if (isHaveWBDS) {
                littleCount += 2;
            }
        }
        column = voteType.length * 2 + littleCount + 4;
        return column;
    }

    // 计算所有分组的项的个数
    public int countRow(String[][] str) {
        int count = 0;
        for (String[] s : str) {
            for (int j = 0; j < s.length; j++) {
                count++;
            }
        }
        return count;
    }

    // 计算每个分组第一行的行数
    public Set<Integer> calculateGroupStartRow(String[][] items, int start) {
        HashSet<Integer> set = new HashSet<>();
        set.add(start);
        for (int i = 0; i < items.length - 1; i++) {
            start += items[i].length;
            set.add(start);
        }
        return set;
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
        cellStyle.setFontFamily("宋体");
        cellStyle.setFontSize(10);
        cellStyle.setColor("000000");
        return cellStyle;
    }

    // 设置 cell 格样式
    public Style getCell9Style() {
        Style cellStyle = new Style();
        cellStyle.setFontFamily("宋体");
        cellStyle.setFontSize(9);
        cellStyle.setColor("000000");
        return cellStyle;
    }

    // 设置 cell 格样式
    public Style getCell8Style() {
        Style cellStyle = new Style();
        cellStyle.setFontFamily("宋体");
        cellStyle.setFontSize(8);
        cellStyle.setColor("000000");
        return cellStyle;
    }

    // 设置 cell 格样式
    public Style getCell6Style() {
        Style cellStyle = new Style();
        cellStyle.setFontFamily("宋体");
        cellStyle.setFontSize(6);
        cellStyle.setColor("000000");
        return cellStyle;
    }

    // 设置 cell 格样式
    public Style getDataCellStyle() {
        Style cellStyle = new Style();
        cellStyle.setFontFamily("宋体");
        cellStyle.setFontSize(8);
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
