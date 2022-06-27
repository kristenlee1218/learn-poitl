package com.learn.test2;

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

    public static final String[] group = new String[]{"政治思想建设", "企业发展质量", "党建工作质量", "作风建设成效"};
    public static final String[] item = new String[]{"政治忠诚", "政治担当", "社会责任", "改革创新", "经营效益", "管理效能", "风险管控", "选人用人", "基层党建", "党风廉政", "团结协作", "联系群众"};
    public static final String[] value1 = new String[]{"101.0", "102.0", "103.0", "104.0", "105.0", "106.0", "107.0", "108.0", "109.0", "110.0", "111.0", "112.0"};
    public static final String[] value2 = new String[]{"111.0", "112.0", "113.0", "114.0", "115.0", "116.0"};
    public static final String[] totalValue = new String[]{"1110.0", "1120.0", "1130.0", "1140.0", "1150.0", "1160.0"};

    @Override
    protected void afterRender(RenderContext<Object> renderContext) {
        // 清空标签
        clearPlaceholder(renderContext, true);
    }

    @Override
    public void doRender(RenderContext<Object> renderContext) throws Exception {
        XWPFRun run = renderContext.getRun();
        // 当前位置的容器
        BodyContainer bodyContainer = BodyContainerFactory.getBodyContainer(run);
        // 定义行列
        int row = 16, col = 14;
        // 当前位置插入表格
        XWPFTable table = bodyContainer.insertNewTable(run, row, col);
        TableTools.widthTable(table, MiniTableRenderData.WIDTH_A4_FULL, 10);
        // 设置表格居中
        TableStyle style = new TableStyle();
        style.setAlign(STJc.CENTER);
        TableTools.styleTable(table, style);

        RowRenderData header0 = RowRenderData.build(new TextRenderData("信息中心领导班子2020年度综合测评汇总表"));
        TableTools.mergeCellsHorizonal(table, 0, 0, col - 1);
        //String[] str = new String[]{"评价", "内部", "外部", "全体"};
        //RowRenderData header1 = RowRenderData.build(str);
        RowRenderData header1 = RowRenderData.build(new TextRenderData("评价内容"), new TextRenderData("内部测评"), new TextRenderData("外部董事"), new TextRenderData("全体"));
        // 分别给以上几个 TextRenderData 设置合并单元格
        TableTools.mergeCellsHorizonal(table, 1, 0, 1);
        TableTools.mergeCellsHorizonal(table, 2, 0, 1);
        TableTools.mergeCellsVertically(table, 11, 1, 2);
        TableTools.mergeCellsVertically(table, 9, 1, 2);
        TableTools.mergeCellsVertically(table, 0, 1, 2);

        TableTools.mergeCellsHorizonal(table, 1, 1, 8);

        TableTools.mergeCellsHorizonal(table, 1, 2, 3);
        TableTools.mergeCellsHorizonal(table, 2, 9, 10);

        TableTools.mergeCellsHorizonal(table, 1, 3, 4);
        TableTools.mergeCellsHorizonal(table, 2, 10, 11);

        RowRenderData header2 = RowRenderData.build(new TextRenderData(""), new TextRenderData("领导班子成员(A1/A2/A3票)"), new TextRenderData("原领导班子成员和中层管理人员(B票)"), new TextRenderData("职工代表(C票"), new TextRenderData("小计"));
        // 分别给以上几个 TextRenderData 设置合并单元格、第一个格子必须写空值否则无法写入到 table
        TableTools.mergeCellsHorizonal(table, 2, 1, 2);
        TableTools.mergeCellsHorizonal(table, 2, 2, 3);
        TableTools.mergeCellsHorizonal(table, 2, 3, 4);
        TableTools.mergeCellsHorizonal(table, 2, 4, 5);

        MiniTableRenderPolicy.Helper.renderRow(table, 0, header0);
        MiniTableRenderPolicy.Helper.renderRow(table, 1, header1);
        MiniTableRenderPolicy.Helper.renderRow(table, 2, header2);

        // group 的设置、集中在第1列
        RowRenderData groupData0 = RowRenderData.build(new TextRenderData(group[0]));
        TableTools.mergeCellsVertically(table, 0, 3, 5);
        TableTools.mergeCellsVertically(table, 0, 6, 9);
        TableTools.mergeCellsVertically(table, 0, 10, 12);
        TableTools.mergeCellsVertically(table, 0, 13, 14);
        MiniTableRenderPolicy.Helper.renderRow(table, 3, groupData0);

        RowRenderData groupData1 = RowRenderData.build(new TextRenderData(group[1]));
        MiniTableRenderPolicy.Helper.renderRow(table, 6, groupData1);

        RowRenderData groupData2 = RowRenderData.build(new TextRenderData(group[2]));
        MiniTableRenderPolicy.Helper.renderRow(table, 10, groupData2);

        RowRenderData groupData3 = RowRenderData.build(new TextRenderData(group[3]));
        MiniTableRenderPolicy.Helper.renderRow(table, 13, groupData3);

        // item 的设置、集中在第 2 列
        RowRenderData itemData0 = RowRenderData.build(new TextRenderData(""), new TextRenderData(item[0]));
        MiniTableRenderPolicy.Helper.renderRow(table, 3, itemData0);

        RowRenderData itemData1 = RowRenderData.build(new TextRenderData(""), new TextRenderData(item[1]));
        MiniTableRenderPolicy.Helper.renderRow(table, 4, itemData1);

        RowRenderData itemData2 = RowRenderData.build(new TextRenderData(""), new TextRenderData(item[2]));
        MiniTableRenderPolicy.Helper.renderRow(table, 5, itemData2);

        RowRenderData itemData3 = RowRenderData.build(new TextRenderData(""), new TextRenderData(item[3]));
        MiniTableRenderPolicy.Helper.renderRow(table, 6, itemData3);

        RowRenderData itemData4 = RowRenderData.build(new TextRenderData(""), new TextRenderData(item[4]));
        MiniTableRenderPolicy.Helper.renderRow(table, 7, itemData4);

        RowRenderData itemData5 = RowRenderData.build(new TextRenderData(""), new TextRenderData(item[5]));
        MiniTableRenderPolicy.Helper.renderRow(table, 8, itemData5);

        RowRenderData itemData6 = RowRenderData.build(new TextRenderData(""), new TextRenderData(item[6]));
        MiniTableRenderPolicy.Helper.renderRow(table, 9, itemData6);

        RowRenderData itemData7 = RowRenderData.build(new TextRenderData(""), new TextRenderData(item[7]));
        MiniTableRenderPolicy.Helper.renderRow(table, 10, itemData7);

        RowRenderData itemData8 = RowRenderData.build(new TextRenderData(""), new TextRenderData(item[8]));
        MiniTableRenderPolicy.Helper.renderRow(table, 11, itemData8);

        RowRenderData itemData9 = RowRenderData.build(new TextRenderData(""), new TextRenderData(item[9]));
        MiniTableRenderPolicy.Helper.renderRow(table, 12, itemData9);

        RowRenderData itemData10 = RowRenderData.build(new TextRenderData(""), new TextRenderData(item[10]));
        MiniTableRenderPolicy.Helper.renderRow(table, 13, itemData10);

        RowRenderData itemData11 = RowRenderData.build(new TextRenderData(""), new TextRenderData(item[11]));
        MiniTableRenderPolicy.Helper.renderRow(table, 14, itemData11);

        // 合并 item 均分的单元格
        TableTools.mergeCellsVertically(table, 3, 3, 5);
        TableTools.mergeCellsVertically(table, 3, 6, 9);
        TableTools.mergeCellsVertically(table, 3, 10, 12);
        TableTools.mergeCellsVertically(table, 3, 13, 14);

        TableTools.mergeCellsVertically(table, 5, 3, 5);
        TableTools.mergeCellsVertically(table, 5, 6, 9);
        TableTools.mergeCellsVertically(table, 5, 10, 12);
        TableTools.mergeCellsVertically(table, 5, 13, 14);

        TableTools.mergeCellsVertically(table, 7, 3, 5);
        TableTools.mergeCellsVertically(table, 7, 6, 9);
        TableTools.mergeCellsVertically(table, 7, 10, 12);
        TableTools.mergeCellsVertically(table, 7, 13, 14);

        TableTools.mergeCellsVertically(table, 9, 3, 5);
        TableTools.mergeCellsVertically(table, 9, 6, 9);
        TableTools.mergeCellsVertically(table, 9, 10, 12);
        TableTools.mergeCellsVertically(table, 9, 13, 14);

        TableTools.mergeCellsVertically(table, 11, 3, 5);
        TableTools.mergeCellsVertically(table, 11, 6, 9);
        TableTools.mergeCellsVertically(table, 11, 10, 12);
        TableTools.mergeCellsVertically(table, 11, 13, 14);

        TableTools.mergeCellsVertically(table, 13, 3, 5);
        TableTools.mergeCellsVertically(table, 13, 6, 9);
        TableTools.mergeCellsVertically(table, 13, 10, 12);
        TableTools.mergeCellsVertically(table, 13, 13, 14);

        // 最后一行的加权总得分
        RowRenderData total = RowRenderData.build(new TextRenderData("加权汇总得分"));
        MiniTableRenderPolicy.Helper.renderRow(table, 15, total);
        TableTools.mergeCellsHorizonal(table, 15, 0, 1);
        TableTools.mergeCellsHorizonal(table, 15, 1, 2);
        TableTools.mergeCellsHorizonal(table, 15, 2, 3);
        TableTools.mergeCellsHorizonal(table, 15, 3, 4);
        TableTools.mergeCellsHorizonal(table, 15, 4, 5);
        TableTools.mergeCellsHorizonal(table, 15, 5, 6);
        TableTools.mergeCellsHorizonal(table, 15, 6, 7);

        RowRenderData lineValue1 = RowRenderData.build(new TextRenderData(""), new TextRenderData(""), new TextRenderData(value1[0]), new TextRenderData(value1[1]), new TextRenderData(value1[2]), new TextRenderData(value1[3]), new TextRenderData(value1[4]), new TextRenderData(value1[5]), new TextRenderData(value1[6]), new TextRenderData(value1[7]), new TextRenderData(value1[8]), new TextRenderData(value1[9]), new TextRenderData(value1[10]), new TextRenderData(value1[11]));
        MiniTableRenderPolicy.Helper.renderRow(table, 3, lineValue1);
        RowRenderData lineValue2 = RowRenderData.build(new TextRenderData(""), new TextRenderData(""), new TextRenderData(value2[0]), new TextRenderData(""), new TextRenderData(value2[1]), new TextRenderData(""), new TextRenderData(value2[2]), new TextRenderData(""), new TextRenderData(value2[3]), new TextRenderData(""), new TextRenderData(value2[4]), new TextRenderData(""), new TextRenderData(value2[5]), new TextRenderData(""));
        MiniTableRenderPolicy.Helper.renderRow(table, 4, lineValue2);
        RowRenderData lineValueLast = RowRenderData.build(new TextRenderData(""), new TextRenderData(totalValue[0]), new TextRenderData(totalValue[1]), new TextRenderData(totalValue[2]), new TextRenderData(totalValue[3]), new TextRenderData(totalValue[4]), new TextRenderData(totalValue[5]));
        MiniTableRenderPolicy.Helper.renderRow(table, 15, lineValueLast);
    }
}
