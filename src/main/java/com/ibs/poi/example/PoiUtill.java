package com.ibs.poi.example;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

/**
 * excel读写工具类
 */
public class PoiUtill {


    private final static String xls = "xls";
    private final static String xlsx = "xlsx";

    /**
     * 读入excel文件，解析后返回
     */
    public static List<String[]> readExcel(File file) throws IOException {

        checkFile(file);
        //获得Workbook工作薄对象
        Workbook workbook = getWorkBook(file);
        //创建返回对象，把每行中的值作为一个数组，所有行作为一个集合返回
        List<String[]> list = new ArrayList<String[]>();
        if (workbook != null) {
            for (int sheetNum = 0; sheetNum < workbook.getNumberOfSheets(); sheetNum++) {
                Sheet sheet = workbook.getSheetAt(sheetNum);
                if (sheet == null) {
                    continue;
                }
                int firstRowNum = sheet.getFirstRowNum();
                int lastRowNum = sheet.getLastRowNum();
                for (int rowNum = firstRowNum + 1; rowNum <= lastRowNum; rowNum++) {
                    Row row = sheet.getRow(rowNum);
                    if (row == null) {
                        continue;
                    }
                    int firstCellNum = row.getFirstCellNum();
                    int lastCellNum = row.getPhysicalNumberOfCells();
                    String[] cells = new String[row.getPhysicalNumberOfCells()];
                    for (int cellNum = firstCellNum; cellNum < lastCellNum; cellNum++) {
                        Cell cell = row.getCell(cellNum);
                        cells[cellNum] = getCellValue(cell);
                    }
                    list.add(cells);
                }
            }
            workbook.close();
        }
        return list;
    }

    public static void checkFile(File file) throws IOException {
        if (null == file) {
            throw new FileNotFoundException("文件不存在！");
        }
        if (file.isDirectory()) {
            throw new FileNotFoundException("该路径为文件夹，请输入文件路径");
        }
        String fileName = file.getName();
        if (!fileName.endsWith(xls) && !fileName.endsWith(xlsx)) {
            throw new IOException(fileName + "不是excel文件");
        }
    }

    public static Workbook getWorkBook(File file) throws IOException {

        String fileName = file.getName();
        Workbook workbook = null;
        InputStream is = new FileInputStream(file);
        if (fileName.endsWith(xls)) {
            //2003
            workbook = new HSSFWorkbook(is);
        } else if (fileName.endsWith(xlsx)) {
            //2007
            workbook = new XSSFWorkbook(is);
        }

        return workbook;
    }

    public static String getCellValue(Cell cell) {
        String cellValue = "";
        if (cell == null) {
            return cellValue;
        }
        //把数字当成String来读，避免出现1读成1.0的情况
        if (cell.getCellType() == CellType.NUMERIC) {
            cell.setCellType(CellType.STRING);
        }
        switch (cell.getCellType()) {
            case NUMERIC:
                cellValue = String.valueOf(cell.getNumericCellValue());
                break;
            case STRING:
                cellValue = String.valueOf(cell.getStringCellValue());
                break;
            case BOOLEAN:
                cellValue = String.valueOf(cell.getBooleanCellValue());
                break;
            case FORMULA:
                cellValue = String.valueOf(cell.getCellFormula());
                break;
            case BLANK:
                cellValue = "";
                break;
            case ERROR: 
                cellValue = "非法字符";
                break;
            default:
                cellValue = "未知类型";
                break;
        }
        return cellValue;
    }
}

