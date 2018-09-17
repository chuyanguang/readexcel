package com.surfront.modules.mms.excel;

//import jxl.Cell;
//import jxl.Sheet;
//import jxl.Workbook;
//import jxl.read.biff.BiffException;


import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;

/**
 * Created by xiong.qiang on 2015/12/9
 */
public class ExcelImporterTester {
    private static DecimalFormat df = new DecimalFormat("0");

    public static void main(String[] args) {
        try {
//            readMyExcel();
            readExcel();

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 读取Excel测试，兼容 Excel 2003/2007/2010
     */
    public static void readExcel() {
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        try {
            //同时支持Excel 2003、2007
            File excelFile = new File("D:\\\\Downloads\\\\test.xls"); //创建文件对象
            FileInputStream is = new FileInputStream(excelFile); //文件流
            Workbook workbook = WorkbookFactory.create(is); //这种方式 Excel 2003/2007/2010 都是可以处理的
            int sheetCount = workbook.getNumberOfSheets();  //Sheet的数量
            //遍历每个Sheet
            for (int s = 0; s < sheetCount; s++) {
                Sheet sheet = workbook.getSheetAt(s);
                int rowCount = sheet.getPhysicalNumberOfRows(); //获取总行数
                //遍历每一行
                for (int r = 0; r < rowCount; r++) {
                    Row row = sheet.getRow(r);
                    int cellCount = row.getPhysicalNumberOfCells(); //获取总列数
                    //遍历每一列
                    for (int c = 0; c < cellCount; c++) {
                        Cell cell = row.getCell(c);
                        int cellType = cell.getCellType();
                        String cellValue = null;
                        switch (cellType) {
                            case Cell.CELL_TYPE_STRING: //文本
                                cellValue = cell.getStringCellValue();
                                break;
                            case Cell.CELL_TYPE_NUMERIC: //数字、日期
                                if (DateUtil.isCellDateFormatted(cell)) {
                                    cellValue = sdf.format(cell.getDateCellValue()); //日期型
                                } else {
                                    cellValue = df.format(cell.getNumericCellValue()); //数字
                                }
                                break;
                            case Cell.CELL_TYPE_BOOLEAN: //布尔型
                                cellValue = String.valueOf(cell.getBooleanCellValue());
                                break;
                            case Cell.CELL_TYPE_BLANK: //空白
                                cellValue = cell.getStringCellValue();
                                break;
                            case Cell.CELL_TYPE_ERROR: //错误
                                cellValue = "";
                                break;
                            case Cell.CELL_TYPE_FORMULA: //公式
                                cellValue = "";
                                break;
                            default:
                                cellValue = "";
                        }
                        System.out.print(cellValue + "    ");
                    }
                    System.out.println();
                }
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static void readMyExcel() throws IOException, InvalidFormatException {
        // jxl 读取 office execl 2003 ，不能解析 office 2007
//            Workbook workbook = Workbook.getWorkbook(new File("D:\\Downloads\\test.xls"));
//            Sheet sheet = workbook.getSheet(0);
//            Cell cell = sheet.getCell(0, 0);
//            String contents = cell.getContents();
//            System.out.println(contents);

        // poi 解析 office execl 2007

        Workbook workbook = WorkbookFactory.create(new FileInputStream(new File("D:\\Downloads\\test.xlsx")));
        Sheet sheet = workbook.getSheetAt(0);
        int firstRowNum = sheet.getFirstRowNum();
        int lastRowNum = sheet.getLastRowNum();
        for (int i = firstRowNum; i <= lastRowNum; i++) {
            Row row = sheet.getRow(i);
            short firstCellNum = row.getFirstCellNum();
            short lastCellNum = row.getLastCellNum();
            for (int j = firstCellNum; j < lastCellNum; j++) {
                Cell cell = row.getCell(j);
                int type = cell.getCellType();
                if (type == Cell.CELL_TYPE_NUMERIC) {
                    String value = df.format(cell.getNumericCellValue());
                    System.out.println(value);
                } else if (type == Cell.CELL_TYPE_STRING) {
                    System.out.println(cell.getStringCellValue());
                } else if (type == Cell.CELL_TYPE_BLANK) {

                }
            }
        }
    }
}
