package com.learn.example2;

import java.util.List;

import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import com.deepoove.poi.data.RowRenderData;
import com.deepoove.poi.policy.DynamicTableRenderPolicy;
import com.deepoove.poi.policy.MiniTableRenderPolicy;
import com.deepoove.poi.util.TableTools;

/**
 * @author ：Kristen
 * @date ：2022/6/13
 * @description : 渲染策略
 */
public class DetailTablePolicy extends DynamicTableRenderPolicy {

    int goodsStartRow = 2;
    int laborsStartRow = 5;

    @Override
    public void render(XWPFTable table, Object data) {
        if (null == data) {
            return;
        }
        DetailData detailData = (DetailData) data;

        List<RowRenderData> labors = detailData.getLabors();
        if (null != labors) {
            table.removeRow(laborsStartRow);
            // 循环插入行
            for (RowRenderData labor : labors) {
                XWPFTableRow insertNewTableRow = table.insertNewTableRow(laborsStartRow);
                for (int j = 0; j < 7; j++) {
                    insertNewTableRow.createCell();
                }

                // 合并单元格
                TableTools.mergeCellsHorizonal(table, laborsStartRow, 0, 3);
                MiniTableRenderPolicy.Helper.renderRow(table, laborsStartRow, labor);
            }
        }

        List<RowRenderData> goods = detailData.getGoods();
        if (null != goods) {
            table.removeRow(goodsStartRow);
            for (RowRenderData good : goods) {
                XWPFTableRow insertNewTableRow = table.insertNewTableRow(goodsStartRow);
                for (int j = 0; j < 7; j++) {
                    insertNewTableRow.createCell();
                }
                MiniTableRenderPolicy.Helper.renderRow(table, goodsStartRow, good);
            }
        }
    }
}