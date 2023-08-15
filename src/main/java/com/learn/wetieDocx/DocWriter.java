package com.learn.wetieDocx;

import org.apache.poi.ooxml.POIXMLDocument;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xwpf.usermodel.*;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class DocWriter {
    public static void writer(String inputSrc, String outSrc, Map<String, String> map) {

        try {
            XWPFDocument doc = new XWPFDocument(POIXMLDocument.openPackage(inputSrc));

            for (XWPFParagraph p : doc.getParagraphs()) {
                List<XWPFRun> runs = p.getRuns();
                if (runs != null) {
                    for (XWPFRun r : runs) {
                        String text = r.getText(0);
                        for (String key : map.keySet()) {
                            if (text != null && text.equals(key)) {
                                r.addBreak();//换行
                                r.setText(map.get(key));
                                r.addBreak();
                                r.setText("############");

                            }
                        }
                    }
                }
            }
            for (XWPFTable tab : doc.getTables()) {
                for (XWPFTableRow row : tab.getRows()) {
                    for (XWPFTableCell cell : row.getTableCells()) {
                        for (XWPFParagraph p : cell.getParagraphs()) {
                            for (XWPFRun r : p.getRuns()) {
                                String text = r.getText(0);
                                for (String key : map.keySet()) {
                                    if (text.equals(key)) {
                                        r.setText(map.get(text), 0);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            doc.write(Files.newOutputStream(Paths.get(outSrc)));
            System.out.println("替换完成");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void main(String[] args) throws IOException, InvalidFormatException {
        Map<String, String> map = new HashMap<String, String>();
        map.put("================", "同意！ CE2988/张三  2019-01-21");
        //文件路径
        String srcPath = "D:\\test-poitl\\3\\1.docx";
        //替换后新文件的路径
        String destPath = "D:\\test-poitl\\3\\1-1.docx";
        writer(srcPath, destPath, map);
    }
}
