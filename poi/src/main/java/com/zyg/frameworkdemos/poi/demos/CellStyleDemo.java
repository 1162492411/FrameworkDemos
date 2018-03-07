package com.zyg.frameworkdemos.poi.demos;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * CellStyle学习示例，CellStyle用于装饰单元格
 */
public class CellStyleDemo {

    /**
     * 为多个单元格使用同一个CellStyle
     * @param file
     */
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
        //背景色和前景色待设置只有设置了fillpattern才有效
        cellStyle.setFillForegroundColor(IndexedColors.GREEN.getIndex());
        cellStyle.setFillBackgroundColor(IndexedColors.ORANGE.getIndex());//仅设置背景色无效，只有在设置类前景色后设置背景色才有效
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

    /**
     * 为单元格设置border
     * @param file
     */
    public static void setCellBorder(File file){
        Workbook workbook = new HSSFWorkbook();
        Sheet sheet = workbook.createSheet("测试单元格边框");
        for (int i = 0; i < 5; i++) {//create al cells
            Row row = sheet.createRow(i);
            for (int j = 0; j < 10; j++) {
                row.createCell(j);
            }
        }
        Cell cellB2 = sheet.getRow(1).getCell(1);
        CellStyle cellStyleB2 = workbook.createCellStyle();
        cellStyleB2.setBorderLeft(BorderStyle.DOUBLE);
        cellStyleB2.setLeftBorderColor(IndexedColors.GREEN.getIndex());
        cellStyleB2.setBorderTop(BorderStyle.THICK);
        cellStyleB2.setTopBorderColor(IndexedColors.BLUE.getIndex());
        cellStyleB2.setBorderRight(BorderStyle.DASH_DOT);
        cellB2.setCellStyle(cellStyleB2);

        Cell cellC3 = sheet.getRow(2).getCell(2);
        CellStyle cellStyleC3 = workbook.createCellStyle();
        cellStyleC3.setBorderTop(BorderStyle.MEDIUM_DASH_DOT);
        cellStyleC3.setBorderBottom(BorderStyle.MEDIUM_DASH_DOT);
        cellStyleC3.setBorderLeft(BorderStyle.MEDIUM_DASH_DOT);
        cellStyleC3.setBorderRight(BorderStyle.MEDIUM_DASH_DOT);
        cellC3.setCellStyle(cellStyleC3);

        try(FileOutputStream stream = new FileOutputStream(file)){
            workbook.write(stream);
        } catch (IOException e) {
            e.printStackTrace();
        }
        System.out.println("setCellBorder done");
    }

    public static void mergedRegion(File file){
        //prepare xls
        Workbook workbook = new HSSFWorkbook();
        Sheet sheet = workbook.createSheet("合并单元格");
        for (int i = 0; i < 10; i++) {
            Row row = sheet.createRow(i);
            for (int j = 0; j < 10; j++) {
                Cell cell = row.createCell(j);
                cell.setCellValue("autoFill-" + i + "-" + j);
            }
        }
        //mergedRegion
        sheet.addMergedRegion(new CellRangeAddress(1,3,2,5));
        Cell cellMergedRegion1 = sheet.getRow(1).getCell(2);
        cellMergedRegion1.setCellStyle(createCenterCellStyle(workbook));

        try(FileOutputStream stream = new FileOutputStream(file)){
            workbook.write(stream);
        } catch (IOException e) {
            e.printStackTrace();
        }
        System.out.println("mergedRegion done");

    }

    private static CellStyle createCenterCellStyle(Workbook workbook){
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        return cellStyle;
    }



    public static void main(String[] args) {
        setSameCellStyle(new File("setSameCellStyle.xls"));
        setCellBorder(new File("setCellBorder.xls"));
        mergedRegion(new File("mergedRegion.xls"));

    }

}
