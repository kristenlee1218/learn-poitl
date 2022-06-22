package com.learn.word;

import org.apache.poi.xwpf.usermodel.*;

import java.util.List;

/**
 * @author ：Kristen
 * @date ：2022/6/22
 * @description : 段落 XWPFParagraph 对象及其操作
 * 段落 XWPFParagraph、段落是构成 Word 文档的一个基本单元
 */
public class Test2 {
    public static void main(String[] args) {
        // 1. 创建新文档
        XWPFDocument doc = new XWPFDocument();
        // 2. 创建新段落
        XWPFParagraph paragraph = doc.createParagraph();

        // 3. 设置段落格式
        // 对齐方式
        paragraph.setAlignment(ParagraphAlignment.CENTER);
        // 边框
        paragraph.setBorderBottom(Borders.DOUBLE);
        paragraph.setBorderTop(Borders.DOUBLE);
        paragraph.setBorderRight(Borders.DOUBLE);
        paragraph.setBorderLeft(Borders.DOUBLE);
        paragraph.setBorderBetween(Borders.SINGLE);

        // 4. 基本元素 XWPFRun
        // XWPFRun 是段落的基本组成单元，它可以是一个文本，也可以是一张图片
        // 段落文本
        // 2.4.1. 读取段落内容
        // 获取文字
        String text = paragraph.getText();

        // 获取段落内所有XWPFRun
        List<XWPFRun> runs = paragraph.getRuns();

        // 创建 XWPFRun 文本
        // 段落末尾创建 XWPFRun
        XWPFRun run = paragraph.createRun();
        run.setText("为这个段落追加文本");


    }
}
