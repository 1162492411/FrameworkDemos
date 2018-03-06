package com.zyg.frameworkdemos.poi.demos;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

/**
 * 读取Excel学习示例
 */
public class ReadXlsDemo {

    private static String BLANK_VALUE = "-";
    private static String ERROR_VALUE = "X";

    //读取一个多Sheet的单元格类型各自不同的xls
    public static void readXlsByCellType(File file) {
        if (file.isFile() && file.exists()) {
            try (FileInputStream stream = new FileInputStream(file)) {
                if (stream.available() > 0) {
                    System.out.println("getXls success " + stream.available());
                    HSSFWorkbook hssfWorkbook = new HSSFWorkbook(stream);
                    System.out.println(file.getName() + "共有表单" + hssfWorkbook.getNumberOfSheets());
                    for (int i = 0; i < hssfWorkbook.getNumberOfSheets(); i++) {
                        HSSFSheet hssfSheet = hssfWorkbook.getSheetAt(i);
                        System.out.println("-----start iterator" + hssfSheet.getSheetName());
                        Iterator<Row> rowIterator = hssfSheet.rowIterator();
                        while (rowIterator.hasNext()) {
                            Row row = rowIterator.next();
                            Iterator<Cell> cellIterator = row.cellIterator();
                            while (cellIterator.hasNext()) {
                                Cell cell = cellIterator.next();
                                readCellByType(cell);
                            }
                            System.out.println();
                        }
                        System.out.println("-----end iterator" + hssfSheet.getSheetName());
                    }
                } else {
                    System.out.println("getXls empty");
                }
            } catch (IOException e) {
                System.out.println("getXls error");
            }
        } else {
            System.out.println("it is not a file.");
        }
    }

    public static void readCellAsString(Cell cell){
        System.out.printf("16%s", cell.getStringCellValue() + "|");
    }

    /**
     * 根据固定的每列格式读取xls
     * @param file 待读取的xls
     */
    public static void readXlsByFixed(File file){
        if (file.isFile() && file.exists()) {
            try (FileInputStream stream = new FileInputStream(file)) {
                if (stream.available() > 0) {
                    System.out.println("getXls success " + stream.available());
                    HSSFWorkbook hssfWorkbook = new HSSFWorkbook(stream);
                    System.out.println(file.getName() + "共有表单" + hssfWorkbook.getNumberOfSheets());
                    for (int i = 0; i < hssfWorkbook.getNumberOfSheets(); i++) {
                        HSSFSheet hssfSheet = hssfWorkbook.getSheetAt(i);
                        System.out.println("-----start iterator" + hssfSheet.getSheetName());
                        Iterator<Row> rowIterator = hssfSheet.rowIterator();
                        while (rowIterator.hasNext()) {
                            Row row = rowIterator.next();
                            Cell cell0 = row.getCell(0);
                            System.out.printf("%16s", cell0.getStringCellValue() + "|");
                            Cell cell1 = row.getCell(1);
                            System.out.printf("%16s", ((int)cell1.getNumericCellValue())  + "|");
                            Cell cell2 = row.getCell(2);
                            System.out.printf("%16s", cell2.getNumericCellValue() + "|");
                            Cell cell3 = row.getCell(3);
                            System.out.printf("%16s", cell3.getBooleanCellValue() + "|");
                            Cell cell4 = row.getCell(4);
                            System.out.printf("%16s", cell4.getDateCellValue() + "|");
                            Cell cell5 = row.getCell(5);
                            System.out.printf("%16s", cell5.getDateCellValue() + "|");
                            System.out.println();
                        }
                        System.out.println("-----end iterator" + hssfSheet.getSheetName());
                    }
                } else {
                    System.out.println("getXls empty");
                }
            } catch (IOException e) {
                System.out.println("getXls error");
            }
        } else {
            System.out.println("it is not a file.");
        }
    }

    /**
     * 根据单元格类型读取单元格内的值,
     * 但是对于数字未处理好，如存入整数时会被视为double存入和取出。由此，既然日期和数字都被视为numeric存入，取出时如何动态处理
     * 还有CellType.ERROR的意义，它是在单元格存入数据时即设置cellType为error，代表某种errorCode还是在读取发生异常时采用getErrorCellValue
     * @param cell 单元格
     */
    public static void readCellByType(Cell cell){
        CellType cellType = cell.getCellTypeEnum();
        Object value;
        if(cellType.equals(CellType.STRING)) {
            System.out.print("get StringCellValue-->");
            value = cell.getStringCellValue();
        }
        else if(cellType.equals(CellType.BOOLEAN)){
            System.out.print("get BooleanCellValue");
            value = cell.getBooleanCellValue();
        }
        else if(cellType.equals(CellType.NUMERIC)){
            System.out.print("get NumericCellValue");
            value = cell.getNumericCellValue();
        }
        else if(cellType.equals(CellType.BLANK)){
            System.out.print("get BlankCellValue");
            value = BLANK_VALUE;
        }
        else if(cellType.equals(CellType.FORMULA)){
            System.out.print("get FormulaCellValue");
            value = cell.getStringCellValue();//formula是指excel公式？
        }
        else {
            System.out.print("get ErrorCellValue");
            value = ERROR_VALUE;
        }
        System.out.printf("%16s", value + "|");
    }

    public static void main(String[] args) {
//        readXlsByCellType(new File("create2.xls"));
        readXlsByFixed(new File("create2.xls"));
    }

}
