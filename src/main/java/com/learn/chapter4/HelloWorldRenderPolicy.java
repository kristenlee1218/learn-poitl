package com.learn.chapter4;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.policy.RenderPolicy;
import com.deepoove.poi.template.ElementTemplate;
import com.deepoove.poi.template.run.RunTemplate;
import org.apache.poi.xwpf.usermodel.XWPFRun;

/**
 * @author ：Kristen
 * @date ：2022/6/7
 * @description :
 */

public class HelloWorldRenderPolicy implements RenderPolicy {
    @Override
    public void render(ElementTemplate eleTemplate, Object data, XWPFTemplate template) {
        XWPFRun run = ((RunTemplate) eleTemplate).getRun();
        String thing = "Hello, world";
        run.setText(thing, 0);
    }
}