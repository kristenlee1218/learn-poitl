package com.learn.statictisDetail;

import com.deepoove.poi.policy.AbstractRenderPolicy;
import com.deepoove.poi.render.RenderContext;
import com.deepoove.poi.xwpf.BodyContainer;
import com.deepoove.poi.xwpf.BodyContainerFactory;
import org.apache.poi.xwpf.usermodel.XWPFRun;

/**
 * @author ：Kristen
 * @date ：2023/8/21
 * @description : 通用 policy，生成 organ 和 leader 的明细表分解，通过 element 来区分 organ 和 leader，
 * 通过票种来确认有几个 table
 */
public class StatisticDetailTablePolicy extends AbstractRenderPolicy<Object> {

//    public static String[] group = new String[]{"政治思想建设", "企业发展质量", "党建工作质量", "作风建设成效"};
//    public static String[][] items = new String[][]{{"政治忠诚", "政治担当", "社会责任"}, {"改革创新", "经营效益"}, {"选人用人", "基层党建", "党风廉政"}, {"团结协作", "联系群众"}};
//    public static String[] item = new String[]{};
//    public static String[] voteType = new String[]{"A1票", "A2票","A3票","A4票","B票", "C票"};

    public static String[] group = new String[]{"对党忠诚", "勇于创新", "治企有方", "兴企有为", "清正廉洁"};
    public static String[][] items = new String[][]{{"政治品质", "政治本领"}, {"创新精神", "创新成果"}, {"经营管理能力", "抓党建强党建能力"}, {"担当作为", "履职绩效"}, {"一岗双责", "廉洁从业"}};
    public static String[] item = new String[]{};
    public static String[] voteType = new String[]{"A1票", "A2票", "A3票", "B票", "C票"};

    // 标题+表头+加权总数所占的固定行数
    int rowBase = 4;
    // “得票” 的固定列数
    int colBase = 1;

    // 计算行
    int row;
    // 计算列
    int col;

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
        row = this.getRow(group);
        System.out.println("row: " + row);

    }

    // 计算行（区分有没有一级指标）
    public int getRow(String[] group) {
        // 有一级指标和二级指标
        if (group.length > 0) {
            row = this.countRow(items) + rowBase;
        } else {
            row = item.length + rowBase;
        }
        return row;
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
