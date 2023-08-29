package org.example.poi;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.junit.Test;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;

public class ExcelWriterTest {
    public static final String PATH = "D:\\Program\\code\\java\\easy-excel\\";


    //HssF类是用来处理03版本的excel
    //XssF类是用来处理07版本的excel
    @Test
    public void testHss() throws IOException {
        //创建工作簿  [03版本]
        Workbook workbook = new HSSFWorkbook();

        //创建工作表
        Sheet sheet = workbook.createSheet("第一个sheet");

        //创建第一行
        Row row1 = sheet.createRow(0);

        //创建单元格 [1:1]
        Cell cell11 = row1.createCell(0);
        //[1:2]
        Cell cell12 = row1.createCell(1);
        cell11.setCellValue("today is 2023 8 29");
        cell12.setCellValue(new DateTime().toString("yyyy-MM-DD HH:mm:ss"));

        //生成一张表[03版]
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "我的文件03.xls");

        //将工作薄从内存写入到文件
        workbook.write(fileOutputStream);
        fileOutputStream.close();
    }

    @Test
    public void testXss() throws IOException {
        //创建工作簿  [03版本]
//        Workbook workbook = new HSSFWorkbook();
        //     [07版本]
        Workbook workbook = new XSSFWorkbook();

        //创建工作表
        Sheet sheet = workbook.createSheet("第一个sheet");

        //创建第一行
        Row row1 = sheet.createRow(0);

        //创建单元格 [1:1]
        Cell cell11 = row1.createCell(0);
        //[1:2]
        Cell cell12 = row1.createCell(1);
        cell11.setCellValue("today is 2023 8 29");
        cell12.setCellValue(new DateTime().toString("yyyy-MM-DD HH:mm:ss"));

        //生成一张表[03版]
//        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "我的文件03.xls");
        //07版
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "我的文件07.xlsx");

        //将工作薄从内存写入到文件
        workbook.write(fileOutputStream);
        fileOutputStream.close();
    }

    //Hss过程中使用缓存，不操作磁盘，速度快，但是不能处理65536行以上的数据
    @Test
    public void testHssBigData() throws IOException {
        long l = System.currentTimeMillis();
        //创建工作簿  [03版本]
        Workbook workbook = new HSSFWorkbook();

        //创建工作表
        Sheet sheet = workbook.createSheet("第一个sheet");
        for (int i = 0; i < 65536; i++) {
            Row row = sheet.createRow(i);
            for (int j = 0; j < 10; j++) {
                Cell cell = row.createCell(j);
                cell.setCellValue(i * 10 + j);
            }
        }

        //生成一张表[03版]
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "我的BigDate文件03.xls");

        //将工作薄从内存写入到文件
        workbook.write(fileOutputStream);
        fileOutputStream.close();

        //706 ms
        System.out.println(System.currentTimeMillis() - l + " ms");
    }


    //Xss速度慢，十分消耗内存，但是可以处理超多65536行以上的数据
    @Test
    public void testXssBigData() throws IOException {
        long l = System.currentTimeMillis();
        //创建工作簿  [07版本]
        Workbook workbook = new XSSFWorkbook();

        //创建工作表
        Sheet sheet = workbook.createSheet("第一个sheet");
        for (int i = 0; i < 65536; i++) {
            Row row = sheet.createRow(i);
            for (int j = 0; j < 10; j++) {
                Cell cell = row.createCell(j);
                cell.setCellValue(i * 10 + j);
            }
        }

        //生成一张表[07版]
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "我的BigDate文件07.xlsx");

        //将工作薄从内存写入到文件
        workbook.write(fileOutputStream);
        fileOutputStream.close();

        //4302 ms
        System.out.println(System.currentTimeMillis() - l + " ms");
    }


    //HXss是对Xss的优化，写入大量数据的术后，写入速度更快，内存占用更少
    //会产生临时文件，需要清理
    @Test
    public void testSXssBigData() throws IOException {
        long l = System.currentTimeMillis();
        //创建工作簿  [07版本升级版]
        Workbook workbook = new SXSSFWorkbook();

        //创建工作表
        Sheet sheet = workbook.createSheet("第一个sheet");
        for (int i = 0; i < 65536; i++) {
            Row row = sheet.createRow(i);
            for (int j = 0; j < 10; j++) {
                Cell cell = row.createCell(j);
                cell.setCellValue(i * 10 + j);
            }
        }

        //生成一张表[07版升级版]
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "我的BigDate文件07S.xls");

        //将工作薄从内存写入到文件
        workbook.write(fileOutputStream);
        fileOutputStream.close();

        //清理临时文件
        ((SXSSFWorkbook)workbook).dispose();

        //1211 ms
        System.out.println(System.currentTimeMillis() - l + " ms");
    }
}