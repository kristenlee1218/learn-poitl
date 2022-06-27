package com.learn.wordTemplate;

import com.deepoove.poi.data.MiniTableRenderData;
import com.deepoove.poi.data.RowRenderData;
import com.deepoove.poi.data.TextRenderData;
import com.deepoove.poi.data.style.TableStyle;
import com.deepoove.poi.policy.AbstractRenderPolicy;
import com.deepoove.poi.policy.MiniTableRenderPolicy;
import com.deepoove.poi.render.RenderContext;
import com.deepoove.poi.util.TableTools;
import com.deepoove.poi.xwpf.BodyContainer;
import com.deepoove.poi.xwpf.BodyContainerFactory;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJc;

/**
 * @author ：Kristen
 * @date ：2022/6/22
 * @description :
 */
public class TestPolicy extends AbstractRenderPolicy<Object> {

    //    第一种测试情况（信息中心_2020年第1批综合测评统计报表(1655863588387)）
//    public static String[] group = new String[]{"政治思想建设", "企业发展质量", "党建工作质量", "作风建设成效"};
//    public static String[][] item = new String[][]{{"政治忠诚", "政治担当", "社会责任"}, {"改革创新", "经营效益", "管理效能", "风险管控"}, {"选人用人", "基层党建", "党风廉政"}, {"团结协作", "联系群众"}};
//    public static String[] value1 = new String[]{"101.0", "102.0", "103.0", "104.0", "105.0", "106.0", "107.0", "108.0", "109.0", "110.0", "111.0", "112.0"};
//    public static String[] value2 = new String[]{"111.0", "112.0", "113.0", "114.0", "115.0", "116.0"};
//    public static String[] totalValue = new String[]{"1110.0", "1120.0", "1130.0", "1140.0", "1150.0", "1160.0"};
//    public static String[] voteType = new String[]{"领导班子成员A1、A2、A3票", "中层测评B票", "职工代表C票", "外部董事A4票"};
//    public static String[] voteTypeGroup = new String[]{"内部测评"};
//    public static String[][] innerEvaluate = new String[][]{{"领导班子成员A1、A2、A3票", "中层测评B票", "职工代表C票"}};
//
//    public static int year = 2022;
//    public static String depart = "信息中心";

//    第二种测试情况（上海大区公司_2021年第1批综合考核评价统计表）
//    public static String[] group = new String[]{"政治思想建设", "企业发展质量", "党建工作质量", "作风建设成效"};
//    public static String[][] item = new String[][]{{"政治忠诚", "政治担当", "社会责任"}, {"改革创新", "经营效益"}, {"选人用人", "基层党建", "党风廉政"}, {"团结协作", "联系群众"}};
//    public static String[] voteType = new String[]{"董事长(A1)", "总经理(A2)", "其他领导(A3)", "本单位领导班子成员(B1、B2)", "中层经理人(C)", "员工(D)"};
//    public static String[] voteTypeGroup = new String[]{"大悦城控股领导", "内部测评"};
//    public static String[][] innerEvaluate = new String[][]{{"董事长(A1)", "总经理(A2)", "其他领导(A3)"}, {"本单位领导班子成员(B1、B2)", "中层经理人(C)", "员工(D)"}};
//
//    public static int year = 2022;
//    public static String depart = "信息中心";

    // 第三种测试情况（上海_2021年第1批综合考核评价统计报表(1646623484403)）
    public static String[] group = new String[]{"政治思想建设", "企业发展质量", "党建工作质量", "作风建设成效"};
    public static String[][] item = new String[][]{{"政治忠诚", "政治担当", "社会责任"}, {"改革创新", "经营效益"}, {"选人用人", "基层党建", "党风廉政"}, {"团结协作", "联系群众"}};
    public static String[] voteType = new String[]{"A票", "B票", "C票"};
    public static String[] voteTypeGroup = new String[]{};
    public static String[][] innerEvaluate = new String[][]{{}};

    public static int year = 2022;
    public static String depart = "信息中心";

    // 计算行
    int row = count(item) + 4;
    // 计算列、先判断有无外部董事
    boolean isHaveWBDS = checkWBDS(voteType);
    int col = this.calculateColumn(voteType, voteTypeGroup, isHaveWBDS);

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

        // 当前位置插入表格
        XWPFTable table = bodyContainer.insertNewTable(run, row, col);
        setTableStyle(table);
        setTableTitle(table);
        setTableHeader(table);
        setGroup(table);
        setItem(table);
        setLastRow(table);

        // 设置计算分数的单元格值
//        RowRenderData lineValue1 = RowRenderData.build(new TextRenderData(""), new TextRenderData(""), new TextRenderData(value1[0]), new TextRenderData(value1[1]), new TextRenderData(value1[2]), new TextRenderData(value1[3]), new TextRenderData(value1[4]), new TextRenderData(value1[5]), new TextRenderData(value1[6]), new TextRenderData(value1[7]), new TextRenderData(value1[8]), new TextRenderData(value1[9]), new TextRenderData(value1[10]), new TextRenderData(value1[11]));
//        MiniTableRenderPolicy.Helper.renderRow(table, 3, lineValue1);
//        RowRenderData lineValue2 = RowRenderData.build(new TextRenderData(""), new TextRenderData(""), new TextRenderData(value2[0]), new TextRenderData(""), new TextRenderData(value2[1]), new TextRenderData(""), new TextRenderData(value2[2]), new TextRenderData(""), new TextRenderData(value2[3]), new TextRenderData(""), new TextRenderData(value2[4]), new TextRenderData(""), new TextRenderData(value2[5]), new TextRenderData(""));
//        MiniTableRenderPolicy.Helper.renderRow(table, 4, lineValue2);
//        RowRenderData lineValueLast = RowRenderData.build(new TextRenderData(""), new TextRenderData(totalValue[0]), new TextRenderData(totalValue[1]), new TextRenderData(totalValue[2]), new TextRenderData(totalValue[3]), new TextRenderData(totalValue[4]), new TextRenderData(totalValue[5]));
//        MiniTableRenderPolicy.Helper.renderRow(table, 15, lineValueLast);
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

    // 整个 table 的样式在此设置
    public void setTableStyle(XWPFTable table) {
        TableTools.widthTable(table, MiniTableRenderData.WIDTH_A4_FULL, 10);
        TableTools.borderTable(table, 4);
        // 设置表格居中
        TableStyle style = new TableStyle();
        style.setAlign(STJc.CENTER);
        TableTools.styleTable(table, style);
    }

    // 设置第一行标题的样式
    public void setTableTitle(XWPFTable table) {
        String title = depart + "领导班子" + year + "年度综合测评汇总表";
        RowRenderData header0 = RowRenderData.build(new TextRenderData(title));
        TableTools.mergeCellsHorizonal(table, 0, 0, col - 1);
        MiniTableRenderPolicy.Helper.renderRow(table, 0, header0);
    }

    // 设置第二、三行标题的样式
    public void setTableHeader(XWPFTable table) {
        // 含有分组
        if ((voteTypeGroup != null && voteTypeGroup.length > 0) && (innerEvaluate != null && innerEvaluate.length > 0)) {
            // 处理第二行，先判断第二行的表头内容
            int length = voteTypeGroup.length + 2;
            // 判断是否包含外部董事
            if (isHaveWBDS) {
                length++;
            }
            String[] strHeader1 = new String[length];
            strHeader1[0] = "评价内容";
            strHeader1[length - 1] = "全体";
            if (isHaveWBDS) {
                strHeader1[length - 2] = "外部董事";
            }
            System.arraycopy(voteTypeGroup, 0, strHeader1, 1, voteTypeGroup.length);
            // 构建第二行
            RowRenderData header1 = RowRenderData.build(strHeader1);
            // 垂直合并"全体"
            TableTools.mergeCellsVertically(table, col - 1, 1, 2);
            TableTools.mergeCellsVertically(table, col - 2, 1, 2);

            // 是否含有外部董事
            if (isHaveWBDS) {
                // 垂直合并“外部董事”
                TableTools.mergeCellsVertically(table, col - 2 - 2, 1, 2);
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
                TableTools.mergeCellsHorizonal(table, 2, col - voteTypeGroup.length - 2 - 1 - 1, col - voteTypeGroup.length - 1 - 1 - 1);
                // 水平合并“全体”
                TableTools.mergeCellsHorizonal(table, 1, voteTypeGroup.length + 2, voteTypeGroup.length + 3);
                TableTools.mergeCellsHorizonal(table, 2, col - voteTypeGroup.length - 2 - 1, col - voteTypeGroup.length - 1 - 1);
            } else {
                TableTools.mergeCellsHorizonal(table, 1, voteTypeGroup.length + 1, voteTypeGroup.length + 2);
                TableTools.mergeCellsHorizonal(table, 2, col - voteTypeGroup.length - 1, col - voteTypeGroup.length);
            }
            MiniTableRenderPolicy.Helper.renderRow(table, 1, header1);

            // 处理分组内的票种和小计
            int total = count(innerEvaluate);
            int lengthItem = (total + voteTypeGroup.length) + 1;
            String[] strHeader2 = new String[lengthItem];
            int index = 1;
            for (String[] strings : innerEvaluate) {
                for (String str : strings) {
                    strHeader2[index++] = str;
                }
                strHeader2[index++] = "小计";
            }
            RowRenderData header2 = RowRenderData.build(strHeader2);
            // 分别给以上几个 TextRenderData 设置合并单元格、第一个格子必须写空值否则无法写入到 table
            for (int i = 0; i < strHeader2.length - 1; i++) {
                TableTools.mergeCellsHorizonal(table, 2, i + 1, i + 2);
            }
            MiniTableRenderPolicy.Helper.renderRow(table, 2, header2);
        } else {
            int length = voteType.length + 2;
            String[] strHeader1 = new String[length];
            strHeader1[0] = "评价内容";
            strHeader1[length - 1] = "全体";
            System.arraycopy(voteType, 0, strHeader1, 1, voteType.length);
            // 构建第二行
            RowRenderData header1 = RowRenderData.build(strHeader1);
            for (int i = 0; i < voteType.length + 2; i++) {
                int j = i + 1;
                TableTools.mergeCellsHorizonal(table, 1, i, j);
                TableTools.mergeCellsHorizonal(table, 2, i, j);
            }

            for (int i = voteType.length + 1; i >= 0; i--) {
                TableTools.mergeCellsVertically(table, i, 1, 2);
            }
            MiniTableRenderPolicy.Helper.renderRow(table, 1, header1);
        }
    }

    // 计算所有分组的项的个数
    public int count(String[][] str) {
        int count = 0;
        for (String[] s : str) {
            for (int j = 0; j < s.length; j++) {
                count++;
            }
        }
        return count;
    }

    // group 的设置、集中在第1列
    public void setGroup(XWPFTable table) {
        int start = 3;
        int end;
        for (int i = 0; i < item.length; i++) {
            end = start + item[i].length;
            TableTools.mergeCellsVertically(table, 0, start, end - 1);
            RowRenderData groupData = RowRenderData.build(new TextRenderData(group[i]));
            MiniTableRenderPolicy.Helper.renderRow(table, start, groupData);
            start = end;
        }
    }

    // item 的设置、集中在第2列
    public void setItem(XWPFTable table) {
        // item 的设置、集中在第 2 列
        // 构建 item 列
        int index = 3;
        for (String[] items : item) {
            for (String s : items) {
                RowRenderData itemData = RowRenderData.build(new TextRenderData(""), new TextRenderData(s));
                MiniTableRenderPolicy.Helper.renderRow(table, index++, itemData);
            }
        }

        // 合并 item 列
        int column = 3;
        for (int i = 0; i < col / 2 - 1; i++) {
            int start = 3;
            int end;
            for (String[] strings : item) {
                end = start + strings.length;
                // 合并 item 均分的单元格
                TableTools.mergeCellsVertically(table, column, start, end - 1);
                start = end;
            }
            column += 2;
        }
    }

    // 设置最后一行
    public void setLastRow(XWPFTable table) {
        // 最后一行的加权总得分
        RowRenderData total = RowRenderData.build(new TextRenderData("加权汇总得分"));
        for (int i = 0; i < col / 2; i++) {
            TableTools.mergeCellsHorizonal(table, row - 1, i, i + 1);
        }
        MiniTableRenderPolicy.Helper.renderRow(table, row - 1, total);
    }

    // 将 table 所需的数据值填充到对应的格子中
    public void setTableData(XWPFTable table) {

    }
}
