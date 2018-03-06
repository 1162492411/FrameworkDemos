package com.zyg.frameworkdemos.poi.demos;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import java.io.*;
import java.time.Instant;
import java.util.Calendar;
import java.util.Date;

/**
 * 创建Excel学习示例
 */
public class CreateExcelDemo {

    /**
     * 创建一个多Sheet的单元格值类型均为String的xls
     * @param file 单元格文件
     */
    public static void createXlsMoreSheetWithAllStringCellType(File file) {
        //创建工作簿
        HSSFWorkbook hssfWorkbook = new HSSFWorkbook();
        //创建工作表
        for (int i = 0; i < 5; i++) {
            HSSFSheet hssfSheet = hssfWorkbook.createSheet("学生基础信息" + i);
            //创建表头与标题
            String[] headers = new String[]{"id", "name", "sex", "class"};//表头

            HSSFRow headerRow = hssfSheet.createRow(0);
            for (int j = 0; j < headers.length; j++) {
                HSSFCell headerCell = headerRow.createCell(j);
                headerCell.setCellValue(headers[j]);
            }
            //添加数据
            for (int j = 1; j <= 10; j++) {
                HSSFRow row = hssfSheet.createRow(j);
                for (int k = 0; k < headers.length; k++) {
                    HSSFCell cell = row.createCell(k);
                    cell.setCellValue(headers[k] + "--" + i + "--" + j);
                }
            }
        }
        //保存为文件
        try (FileOutputStream fileOutputStream = new FileOutputStream(file)) {
            hssfWorkbook.write(fileOutputStream);
        } catch (IOException e) {
            e.printStackTrace();
        }
        System.out.println("done");
    }

    /**
     * 创建一个单sheet单row的表格，该表格每列的类型均不同
     * @param file 单元格文件
     */
    public static void createXlsSingleRowWithDifferentCellType(File file) {
        //创建工作簿
        HSSFWorkbook hssfWorkbook = new HSSFWorkbook();
        //创建工作表
            HSSFSheet hssfSheet = hssfWorkbook.createSheet("测试单元格类型");
            //创建表头与标题
            HSSFRow row = hssfSheet.createRow(0);

            Cell cellString = row.createCell(0);
            cellString.setCellValue("stringValue1");
            Cell cellDoubleA = row.createCell(1);
            cellDoubleA.setCellValue(1);
            Cell cellDoubleB = row.createCell(2);
            cellDoubleB.setCellValue(1.11);
            Cell cellBoolean = row.createCell(3);
            cellBoolean.setCellValue(true);
            Cell cellDate = row.createCell(4);
            Date date = Date.from(Instant.now());
            System.out.println(date.getTime());
            cellDate.setCellValue(date);//插入日期时总保存为一个浮点数，且值与日期待毫秒值相差很大
            Cell cellCalender = row.createCell(5);
            cellCalender.setCellValue(Calendar.getInstance());
            Cell cellEmpty = row.createCell(6);

        //保存为文件
        try (FileOutputStream fileOutputStream = new FileOutputStream(file)) {
            hssfWorkbook.write(fileOutputStream);
        } catch (IOException e) {
            e.printStackTrace();
        }
        System.out.println("done");
    }

    public static void main(String[] args) {
        createXlsMoreSheetWithAllStringCellType(new File("create1.xls"));
        createXlsSingleRowWithDifferentCellType(new File("create2.xls"));
    }
}
