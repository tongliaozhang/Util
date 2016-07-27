package com.xbniao.util;

import org.apache.poi.hssf.usermodel.examples.CellTypes;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;

/**
 * Created by zhangql on 2016/7/11.
 */
public class ExcelToSql {
    public static void main(String[] args) throws Exception{
        String fileName = "E:\\hwc.xlsx";
        File f  = new File(fileName);

        String[] fields = {"商户/科目", "商户号/科目号", "资金属性","交易类型", "原子交易类型","借贷","金额","备注"};
        InputStream inputStream = new FileInputStream(f);

        XSSFWorkbook xwb = new XSSFWorkbook(inputStream);
        XSSFSheet sheet = xwb.getSheetAt(0);
        XSSFRow row = null;
        XSSFCell cell = null;
        System.out.println("文件一共"+sheet.getLastRowNum()+"行");
        System.out.println("文件一共"+sheet.getPhysicalNumberOfRows()+"行getPhysicalNumberOfRows");

        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            row = sheet.getRow(i);
            StringBuffer buffer = new StringBuffer();
            for(int j=0;j<17;j++){
                cell = row.getCell(j);
                if(null !=cell){
                    if(XSSFCell.CELL_TYPE_STRING == cell.getCellType()){
                        buffer.append(cell.getStringCellValue()).append("|");
                    }else if(XSSFCell.CELL_TYPE_NUMERIC == cell.getCellType()){
                        buffer.append(cell.getNumericCellValue()).append("|");
                    }else if(XSSFCell.CELL_TYPE_BLANK == cell.getCellType()){
                        buffer.append("#######").append("|");
                    }
                }

            }
            System.out.println(buffer.toString());
        }


    }
}
