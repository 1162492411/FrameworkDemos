package com.zyg.frameworkdemos.poi.demos;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * 公式使用示例
 */
public class FormulaDemo {

    public static void simpleFormulas(File file){
        //prepare xls
        Workbook workbook = new HSSFWorkbook();
        Sheet sheet = workbook.createSheet("测试简单公式");
        for (int i = 0; i < 10; i++) {
            Row row = sheet.createRow(i);
            for (int j = 0; j < 10; j++) {
                row.createCell(j);
            }
        }
        Cell cellNumberA = sheet.getRow(0).getCell(1);
        cellNumberA.setCellValue("NumberA : ");
        Cell cellValueA = sheet.getRow(0).getCell(2);
        cellValueA.setCellValue(3);
        Cell cellNumberB = sheet.getRow(1).getCell(1);
        cellNumberB.setCellValue("NumberB : ");
        Cell cellValueB = sheet.getRow(1).getCell(2);
        cellValueB.setCellValue(7);
        Cell cellFormulaA = sheet.getRow(2).getCell(1);
        cellFormulaA.setCellValue("Total : ");
        Cell cellFormulaValueA = sheet.getRow(2).getCell(2);
//        cellFormulaValueA.setCellType(CellType.FORMULA);
        cellFormulaValueA.setCellFormula("SUM(C1:C2)");
        Cell cellFormulaB = sheet.getRow(3).getCell(1);
        cellFormulaB.setCellValue("MAX : ");
        Cell cellFormulaValueB = sheet.getRow(3).getCell(2);
        cellFormulaValueB.setCellFormula("MAX(C1:C2)");

        try(FileOutputStream stream = new FileOutputStream(file)){
            workbook.write(stream);
        } catch (IOException e) {
            e.printStackTrace();
        }
        System.out.println("simpleFormulas done");
    }


    public static void main(String[] args) {
        simpleFormulas(new File("simpleFormulas.xls"));

    }

}
