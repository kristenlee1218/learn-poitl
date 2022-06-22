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

        RowRenderData header1 = RowRenderData.build(new TextRenderData("评价内容"), new TextRenderData("内部测评"), new TextRenderData("外部董事"), new TextRenderData("全体"));
        TableTools.mergeCellsHorizonal(table, 1, 0, 1);
        TableTools.mergeCellsHorizonal(table, 2, 0, 1);
       // TableTools.mergeCellsVertically(table, 0, 1, 2);

        TableTools.mergeCellsHorizonal(table, 1, 1, 8);

        TableTools.mergeCellsHorizonal(table, 1, 2, 3);
        TableTools.mergeCellsHorizonal(table, 2, 9, 10);
        //TableTools.mergeCellsVertically(table, 2, 1, 2);

        TableTools.mergeCellsHorizonal(table, 1, 3, 4);
        TableTools.mergeCellsHorizonal(table, 2, 10, 11);

        RowRenderData header2 = RowRenderData.build(new TextRenderData("领导班子成员(A1/A2/A3票)"), new TextRenderData("原领导班子成员和中层管理人员(B票)"), new TextRenderData("职工代表(C票"), new TextRenderData("小计"));

        MiniTableRenderPolicy.Helper.renderRow(table, 0, header0);
        MiniTableRenderPolicy.Helper.renderRow(table, 1, header1);

    }
}
