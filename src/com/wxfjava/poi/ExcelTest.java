package com.wxfjava.poi;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.hssf.record.crypto.Biff8EncryptionKey;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * 查找某个excel文件的密码知道组成结构
 * @author wxf
 *
 */
public class ExcelTest {
    
    public static void main(String[] args) {
        for (int i=100; i<1000; i++){
            try {
                readFile("abc"+i);
            } catch (Exception e) {
                
            }
        }
    }

    private static void readFile(String password) throws IOException, FileNotFoundException {
        System.out.println(password);
        Biff8EncryptionKey.setCurrentUserPassword(password);
        Workbook workbook = new HSSFWorkbook(new FileInputStream("/xxxxxxxxxx/a.xls")); 
        
        Sheet sheet = workbook.getSheetAt(0);
        Row row = sheet.getRow(0);
        Cell cell = row.getCell(1);
        String stringCellValue = cell.getStringCellValue();
        System.out.println("==================>"+stringCellValue);
        System.out.println(password);
    }

}
