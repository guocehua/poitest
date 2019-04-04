package com.ibs.poi.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class Poitest {
    private static XSSFWorkbook XSSFWorkbook;

    public static void main(String[] args) throws IOException {
        XSSFWorkbook=new XSSFWorkbook(new FileInputStream("C:\\文档\\test.xlsx")) ;
        Sheet sheet=XSSFWorkbook.getSheetAt(0);
        int start = sheet.getFirstRowNum();
        int end =sheet.getLastRowNum();
        System.out.println("班级  姓名");
        for (int i= start+1;i<end+1;i++){
            Row row= sheet.getRow(i);
            int first= row.getFirstCellNum();
            int last= row.getLastCellNum();
            for(int j=first;j<last+1;j++){
                Cell cell=row.getCell(j);
                if(cell==null)
                    continue;
                switch (cell.getCellType()){
                    case NUMERIC:
                        System.out.print(cell.getNumericCellValue()+"");
                        break;
                    case STRING:
                        System.out.print(cell.getStringCellValue());
                        break;
                }
                System.out.print("  ");
        }
            System.out.println();
    }


}
}
