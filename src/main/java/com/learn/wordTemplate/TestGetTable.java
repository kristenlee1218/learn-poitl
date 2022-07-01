package com.learn.wordTemplate;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.List;

/**
 * @author ：Kristen
 * @date ：2022/7/1
 * @description :
 */
public class TestGetTable {
    public static void main(String[] args) throws IOException {
        XWPFDocument doc = new XWPFDocument(Files.newInputStream(Paths.get("D:\\test-poitl\\ele02_out1.docx")));
        List<XWPFTable> tables = doc.getTables();
        XWPFTable table = tables.get(0);
        for (int i = 3; i < table.getNumberOfRows(); i++) {
            //System.out.println("table.getRow(i).getTableCells().size(): " + table.getRow(i).getTableCells().size());
            for (int j = 2; j < table.getRow(i).getTableCells().size(); j++) {
                System.out.println("table.getRow(i).getCell(j).getText(): " + table.getRow(i).getCell(j).getText());
                table.getRow(i).getCell(j).setText("A");
            }
            System.out.println("================");
        }
    }
}