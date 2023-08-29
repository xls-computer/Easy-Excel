package org.example.poi;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.junit.Test;

import javax.sound.midi.Soundbank;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class ExcelReaderTest {
    public static final String PATH = "D:\\Program\\code\\java\\easy-excel\\";

    //Hss
    @Test
    public void testRead03() throws IOException {
        FileInputStream fileInputStream = new FileInputStream(PATH + "我的文件03.xls");
        //工作簿，使用excel能操作的这里也都能操作
        Workbook workbook = new HSSFWorkbook(fileInputStream);

        //得到表
        Sheet sheet = workbook.getSheetAt(0);
        //得到行
        Row row = sheet.getRow(0);
        //得到单元格
        Cell cell = row.getCell(0);
        //获取String类型的值【读取的时候要注意类型】
        System.out.println(cell.getStringCellValue());
        fileInputStream.close();
    }

    //Xss
    @Test
    public void testRead07() throws IOException {
        FileInputStream fileInputStream = new FileInputStream(PATH + "我的文件03.xls");
        //工作簿，使用excel能操作的这里也都能操作
        Workbook workbook = new HSSFWorkbook(fileInputStream);

        //得到表
        Sheet sheet = workbook.getSheetAt(0);
        //得到行
        Row row = sheet.getRow(0);
        //得到单元格
        Cell cell = row.getCell(0);
        //获取String类型的值【读取的时候要注意类型】
        System.out.println(cell.getStringCellValue());
        fileInputStream.close();
    }


    //读取不同类型的值
    @Test
    public void testReadDiffDataType() throws IOException {
        FileInputStream fileInputStream = new FileInputStream(PATH + "test_data.xls");
        //工作簿，使用excel能操作的这里也都能操作
        Workbook workbook = new HSSFWorkbook(fileInputStream);

        //得到表
        Sheet sheet = workbook.getSheetAt(0);
        Row rowTitle = sheet.getRow(0);
        //获取该行有多少数据
        int rowTitleCount = rowTitle.getPhysicalNumberOfCells();
        for (int i = 0; i < rowTitleCount; i++) {
            Cell cell = rowTitle.getCell(i);
            System.out.print(cell.getStringCellValue() +" |");
        }
        System.out.println();
        //获取有多少行
        int rowsCount = sheet.getPhysicalNumberOfRows();
        for (int i = 1; i < rowsCount; i++) {
            Row rowData = sheet.getRow(i);
            int cellCount = rowData.getPhysicalNumberOfCells();
            for (int j = 0; j < cellCount; j++) {
                Cell cell = rowData.getCell(j);
                System.out.print(" [" + i  +" "+ (j+1) +"] ");
                if(cell!=null){
                    int cellType = cell.getCellType();
                    switch (cellType){
                        case Cell.CELL_TYPE_BLANK:
                            System.out.println("[BLAND]");
                            break;
                        case Cell.CELL_TYPE_BOOLEAN:
                            System.out.println("[BOOLEAN] " +cell.getBooleanCellValue());
                            break;
                        case Cell.CELL_TYPE_ERROR:
                            System.out.println("[ERROR]");
                            break;
                        case Cell.CELL_TYPE_FORMULA:
                            System.out.println("[FORMULA] " +cell.getCellFormula());
                            FormulaEvaluator formulaEvaluator = new HSSFFormulaEvaluator((HSSFWorkbook) workbook);
                            CellValue evaluate = formulaEvaluator.evaluate(cell);
                            String cellvalue = evaluate.formatAsString();
                            System.out.println("计算结果 "+cellvalue);
                            break;
                        case Cell.CELL_TYPE_STRING:
                            System.out.println("[STRING] " +cell.getStringCellValue());
                            break;
                        case Cell.CELL_TYPE_NUMERIC:
                            System.out.println("[NUMERIC]" + cell.getNumericCellValue());
                    }
                }
            }
            System.out.println();
        }
        fileInputStream.close();
    }




}
