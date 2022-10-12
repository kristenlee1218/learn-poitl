package com.learn.example2;

import java.util.Arrays;
import java.util.Collections;
import java.util.List;

import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJc;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import com.deepoove.poi.data.MiniTableRenderData;
import com.deepoove.poi.data.RowRenderData;
import com.deepoove.poi.data.TextRenderData;
import com.deepoove.poi.data.style.Style;
import com.deepoove.poi.data.style.TableStyle;

/**
 * @author ：Kristen
 * @date ：2022/6/13
 * @description : 付款说明书
 */
@DisplayName("Example for Table")
public class PaymentExampleTest {

    String resource = "D:\\test-poitl\\template2.docx";
    PaymentData data = new PaymentData();

    @BeforeEach
    public void init() {
        Style headTextStyle = new Style();

        headTextStyle.setFontFamily("Hei");
        headTextStyle.setFontSize(9);
        headTextStyle.setColor("7F7F7F");

        TableStyle headStyle = new TableStyle();
        headStyle.setBackgroundColor("F2F2F2");
        headStyle.setAlign(STJc.CENTER);

        TableStyle rowStyle = new TableStyle();
        rowStyle.setAlign(STJc.CENTER);

        data.setNO("KB.6890451");
        data.setID("ZHANG_SAN_091");
        data.setTaitou("深圳XX家装有限公司");
        data.setConsignee("丙丁");

        data.setSubtotal("8000");
        data.setTax("600");
        data.setTransform("120");
        data.setOther("250");
        data.setUnpay("6600");
        data.setTotal("总共：7200");

        RowRenderData header = RowRenderData.build(new TextRenderData("日期", headTextStyle), new TextRenderData("订单编号", headTextStyle), new TextRenderData("销售代表", headTextStyle), new TextRenderData("离岸价", headTextStyle), new TextRenderData("发货方式", headTextStyle), new TextRenderData("条款", headTextStyle), new TextRenderData("税号", headTextStyle));
        header.setRowStyle(headStyle);

        RowRenderData row = RowRenderData.build("2018-06-12", "SN18090", "李四", "5000元", "快递", "附录A", "T11090");
        row.setRowStyle(rowStyle);
        MiniTableRenderData miniTableRenderData = new MiniTableRenderData(header, Collections.singletonList(row), MiniTableRenderData.WIDTH_A4_MEDIUM_FULL);
        miniTableRenderData.setStyle(headStyle);
        data.setOrder(miniTableRenderData);

        DetailData detailTable = new DetailData();
        RowRenderData good = RowRenderData.build("4", "墙纸", "书房+卧室", "1500", "/", "400", "1600");
        good.setRowStyle(rowStyle);
        List<RowRenderData> goods = Arrays.asList(good, good, good);
        RowRenderData labor = RowRenderData.build("油漆工", "2", "200", "400");
        labor.setRowStyle(rowStyle);
        List<RowRenderData> labors = Arrays.asList(labor, labor, labor, labor);
        detailTable.setGoods(goods);
        detailTable.setLabors(labors);
        data.setDetailTable(detailTable);
    }

    @Test
    public void testPaymentExample() throws Exception {
        Configure config = Configure.newBuilder().bind("detail_table", new DetailTablePolicy()).build();
        XWPFTemplate template = XWPFTemplate.compile(resource, config).render(data);
        template.writeToFile("D:\\test-poitl\\template2_out.docx");
    }
}