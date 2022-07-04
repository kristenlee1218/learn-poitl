package com.learn.wordTemplate;

import com.deepoove.poi.policy.AbstractRenderPolicy;
import com.deepoove.poi.render.RenderContext;
import com.deepoove.poi.xwpf.BodyContainer;
import com.deepoove.poi.xwpf.BodyContainerFactory;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;

/**
 * @author ：Kristen
 * @date ：2022/7/4
 * @description :
 */
public class ItemPolicy extends AbstractRenderPolicy<Object> {

    // 第四种测试情况（河北建投二级单位_2022年第1批综合考核评价统计报表(1653979325662)）
    public static String[][] item = new String[][]{{"政治忠诚", "政治担当", "社会责任"}, {"改革创新", "经营效益"}, {"选人用人", "基层党建", "党风廉政"}, {"团结协作", "联系群众"}};
    public static String[] voteType = new String[]{"A票", "B票"};
    public static String[] value = new String[]{"1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31", "32", "33", "34", "35", "36", "37", "38", "39", "40", "41", "42", "43", "44", "45", "46", "47", "48", "49", "50", "51", "52", "53", "54", "55", "56", "57", "58", "59", "60"};
    public static int year = 2022;
    public static String depart = "信息中心";

    // 计算行
    int row = countRow(item) + 4;
    // 计算列
    int col = voteType.length + 2;

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
        System.out.println("row: " + row);
        System.out.println("col: " + col);
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
}
