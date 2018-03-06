package com.zyg.frameworkdemos.poi.demos;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * CellStyle学习示例，CellStyle用于装饰单元格
 */
public class CellStyleDemo {

    public static void setSameCellStyle(File file){
        Workbook workbook = new HSSFWorkbook();
        CellStyle cellStyle1 = createCellStyle1(workbook); //create cellstyle
        Sheet sheet = workbook.createSheet("测试cellStyle1");//create sheet
//        sheet.setDefaultColumnStyle(2,cellStyle1);//set cellstyle for pnt column
        //fill cell value
        for (int i = 0; i < 10; i++) {
            Row row = sheet.createRow(i);
            if(i == 3){
//                row.setRowStyle(cellStyle1);//调用Row.setRowStyle()为某行设置CellStyle,但调用后该行仅无值字段显示了背景色
            }
            row.setHeight((short)500);
            for (int j = 0; j < 5; j++) {
                Cell cell = row.createCell(j);
                cell.setCellValue("autoFill-" + i + "-" + j);
                cell.setCellStyle(cellStyle1);//调用Cell.setCellStyle()为单个的单元格设置CellStyle
            }
        }
        //save xls
        try(FileOutputStream stream = new FileOutputStream(file)) {
            workbook.write(stream);
        } catch (IOException e) {
            e.printStackTrace();
        }
        System.out.println("done");
    }

    /**
     * 生成cellStyle
     * @param workbook
     * @return
     */
    private static CellStyle createCellStyle1(Workbook workbook){
        CellStyle cellStyle  = workbook.createCellStyle();
//        cellStyle.setFillForegroundColor(IndexedColors.GREEN.getIndex());
//        cellStyle.setFillBackgroundColor(IndexedColors.ORANGE.getIndex());//仅设置背景色无效，只有在设置类前景色后设置背景色才有效
        cellStyle.setFillPattern(FillPatternType.LESS_DOTS);//设置FillPattern后单元格颜色背景色为灰色
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.TOP);
        Font font = workbook.createFont();
        font.setFontHeightInPoints((short) 10);
        font.setColor(IndexedColors.GREEN.getIndex());
        font.setFontName("宋体");
        cellStyle.setFont(font);
        return cellStyle;
    }

    public static void main(String[] args) {
        setSameCellStyle(new File("cellStyle1.xls"));
    }

}
