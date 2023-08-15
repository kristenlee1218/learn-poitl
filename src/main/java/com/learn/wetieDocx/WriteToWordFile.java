package com.learn.wetieDocx;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.util.Arrays;

public class WriteToWordFile {
    public static void main(String[] args) {
        try {
            // 打开一个本地 Word 文件
            XWPFDocument document = new XWPFDocument(Files.newInputStream(new File("D:\\test-poitl\\3\\1.docx").toPath()));
            XWPFParagraph para = document.createParagraph();
            XWPFRun run = para.createRun();
            FileOutputStream out = new FileOutputStream("D:\\test-poitl\\3\\1-1.docx");
            String[] str = new String[]{"A1", "A2", "A3", "B", "C"};
            run.setText("aa");
            run.addBreak();
            document.write(out);
            run.setText("bb");
            run.addBreak();
            document.write(out);
            out.close();
            System.out.println("写入成功");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}