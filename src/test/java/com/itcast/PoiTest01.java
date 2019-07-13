package com.itcast;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class PoiTest01 {

    @Test
    public void test() throws IOException {
        //获取要读取的工作册
        Workbook xk = new XSSFWorkbook("D:\\每日课堂\\day10\\03-资料\\poi资料\\demo.xlsx");
        Sheet sheet = xk.getSheetAt(0);
        //getLastRowNum 获取最后一行下标
        for (int i = 0; i < sheet.getLastRowNum()+1; i++) {
            Row row = sheet.getRow(i);
            String str = "";
            //获取最后一列下标
            for (int j = 0; j < row.getLastCellNum(); j++) {
                Cell cell = row.getCell(j);
                if (cell != null) {
                    str += getCellValue(cell)+"\t";
                }


            }
            System.out.println(str);
            System.out.println("我是一个英雄");
            System.out.println("我是大哥");
              

        }


    }


    private Object getCellValue(Cell cell) {

        Object o = null;
        //获取cell值的属性 看走哪一个
        CellType cellType = cell.getCellType();
        switch (cellType) {
            case STRING: //字符串
                o = cell.getStringCellValue();
                break;
            case NUMERIC: //数值
                if(DateUtil.isCellDateFormatted(cell)){
                    o = cell.getDateCellValue().toLocaleString();// yyyy-MM-dd hh:MM:ss
                }else{
                    o = cell.getNumericCellValue();
                }
                break;
            case BOOLEAN: //bool
                o = cell.getBooleanCellValue();
                break;
            default:
                break;
        }
        return o;


    }
}
