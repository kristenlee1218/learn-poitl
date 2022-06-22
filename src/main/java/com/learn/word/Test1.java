package com.learn.word;

import org.apache.poi.xwpf.usermodel.*;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.List;

/**
 * @author ：Kristen
 * @date ：2022/6/22
 * @description : 文档 XWPFDocument 对象及其操作
 * 文档 XWPFDocument
 * XWPFDocument是对 .docx 文档操作的高级封装 API。
 */
public class Test1 {
    public static void main(String[] args) throws IOException {
        // 1. 创建新文档
        XWPFDocument doc1 = new XWPFDocument();

        // 3. 生成文档、通过传入一个OutStream，将文档写入流：
        try (FileOutputStream out = new FileOutputStream("simple.docx")) {
            doc1.write(out);
        }

        // 2. 读取已有文档：段落、表格、图片
        XWPFDocument doc2 = new XWPFDocument(Files.newInputStream(Paths.get("./deepoove.docx")));
        // 段落
        List<XWPFParagraph> paragraphs = doc2.getParagraphs();
        // 表格
        List<XWPFTable> tables = doc2.getTables();
        // 图片
        List<XWPFPictureData> allPictures = doc2.getAllPictures();
        // 页眉
        List<XWPFHeader> headerList = doc2.getHeaderList();
        // 页脚
        List<XWPFFooter> footerList = doc2.getFooterList();
    }
}
