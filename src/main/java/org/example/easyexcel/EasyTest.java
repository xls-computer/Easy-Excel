package org.example.easyexcel;

import com.alibaba.excel.EasyExcel;
import org.junit.Test;

import java.util.ArrayList;
import java.util.Date;

public class EasyTest {
    public static final String PATH = "D:\\Program\\code\\java\\easy-excel\\";

    public ArrayList<DemoData> generateData(){
        ArrayList<DemoData> list = new ArrayList<>(10);
        for (int i = 0; i < 10; i++) {
            DemoData demoData = new DemoData();
            demoData.setString("my "+i);
            demoData.setDate(new Date());
            demoData.setDoubleData((double) System.currentTimeMillis());
            list.add(demoData);
        }
        return list;
    }

    @Test
    public void simpleWrite(){
        String fileName = PATH+ "easy.xlsx";
        //需要指出那个class去写，然后写到第一个sheet，名字为“模板”,然后文件流会自动关闭
        //DemoData.class为格式类
        //sheet(表名)
        //doWrite(数据)
        EasyExcel.write(fileName, DemoData.class).sheet("模板").doWrite(generateData());
    }

    @Test
    public void simpleReader(){
        String fileName = PATH+ "easy.xlsx";

        EasyExcel.read(fileName, DemoData.class, new DemoDataListener()).sheet().doRead();
    }








}
