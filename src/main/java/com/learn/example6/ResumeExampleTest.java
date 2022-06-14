package com.learn.example6;

import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.data.DocxRenderData;
import com.deepoove.poi.data.NumbericRenderData;
import com.deepoove.poi.data.PictureRenderData;
import com.deepoove.poi.data.TextRenderData;
import com.deepoove.poi.data.style.Style;

/**
 * @author ：Kristen
 * @date ：2022/6/14
 * @description :
 */
@DisplayName("Example for Resume")
public class ResumeExampleTest {
    ResumeData data = new ResumeData();

    @BeforeEach
    public void init() {
        data.setPortrait(new PictureRenderData(100, 100, "D:\\test-poitl\\sayi.png"));
        data.setName("卅一");
        data.setJob("BUG工程师");
        data.setPhone("18080809090");
        data.setSex("男");
        data.setProvince("杭州");
        data.setBirthday("2000.08.19");
        data.setEmail("adasai90@gmail.com");
        data.setAddress("浙江省杭州市西湖区古荡1号");
        data.setEnglish("CET6 620");
        data.setUniversity("美国斯坦福大学");
        data.setProfession("生命工程");
        data.setEducation("学士");
        data.setRank("班级排名 1/36");
        data.setHobbies("音乐、画画、乒乓球、旅游、读书\nhttps://github.com/Sayi");

        // 技术栈部分
        TextRenderData textRenderData = new TextRenderData("SpringBoot、SprigCloud、Mybatis");
        Style style = new Style();
        style.setFontSize(10);
        style.setColor("7F7F7F");
        style.setFontFamily("微软雅黑");
        textRenderData.setStyle(style);
        data.setStack(new NumbericRenderData(Arrays.asList(textRenderData, textRenderData, textRenderData)));

        // 模板文档循环合并
        List<ExperienceData> experiences = new ArrayList<>();
        ExperienceData data0 = new ExperienceData();
        data0.setCompany("香港微软");
        data0.setDepartment("Office 事业部");
        data0.setTime("2001.07-2020.06");
        data0.setPosition("BUG工程师");
        textRenderData = new TextRenderData("负责生产BUG，然后修复BUG，同时有效实施招聘行为");
        textRenderData.setStyle(style);
        data0.setResponsibility(new NumbericRenderData(NumbericRenderData.FMT_LOWER_ROMAN, Arrays.asList(textRenderData, textRenderData)));
        ExperienceData data1 = new ExperienceData();
        data1.setCompany("自由职业");
        data1.setDepartment("OpenSource 项目组");
        data1.setTime("2015.07-2020.06");
        data1.setPosition("研发工程师");
        textRenderData = new TextRenderData("开源项目的迭代和维护工作");
        textRenderData.setStyle(style);
        TextRenderData textRenderData1 = new TextRenderData("持续集成、Swagger文档等工具调研");
        textRenderData1.setStyle(style);
        data1.setResponsibility(new NumbericRenderData(NumbericRenderData.FMT_LOWER_ROMAN, Arrays.asList(textRenderData, textRenderData1, textRenderData)));
        experiences.add(data0);
        experiences.add(data1);
        experiences.add(data0);
        experiences.add(data1);
        data.setExperience(new DocxRenderData(new File("D:\\test-poitl\\segment.docx"), experiences));
    }

    @Test
    public void testResumeExample() throws Exception {
        XWPFTemplate template = XWPFTemplate.compile("D:\\test-poitl\\resume.docx").render(data);
        FileOutputStream out = new FileOutputStream("D:\\test-poitl\\resume_out.docx");
        template.write(out);
        out.flush();
        out.close();
        template.close();
    }
}