package com.learn.test2;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * @author ：Kristen
 * @date ：2024/7/26
 * @description :
 */
public class Test2 {
    public static void main(String[] args) {
        // 创建 workbook XSSFWorkbook生成xlsx、HSSFWorkbook 生成 xls
        HSSFWorkbook workbook = new HSSFWorkbook();
        // 创建 sheet
        HSSFSheet sheet = workbook.createSheet("sheet1");
        // 创建样式
        HSSFCellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setShrinkToFit(true);
    }
}
