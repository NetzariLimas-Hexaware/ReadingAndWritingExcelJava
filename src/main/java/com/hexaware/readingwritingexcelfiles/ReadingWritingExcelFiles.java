/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Project/Maven2/JavaApp/src/main/java/${packagePath}/${mainClassName}.java to edit this template
 */

package com.hexaware.readingwritingexcelfiles;

import java.io.File;
import java.io.FileOutputStream;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


/**
 *
 * @author Netzari Limas
 */
public class ReadingWritingExcelFiles {

    public static void main(String[] args) {
        // Blank workbook
        XSSFWorkbook workbook = new XSSFWorkbook();
        
        // Crear un blank sheet
        XSSFSheet sheet = workbook.createSheet();
        
        // This data needs to be writter (Object[])
        Map<String, Object[]> data = new TreeMap<String, Object[]>();
        data.put("1", new Object[] {"NAME", "LASTNAME", "EMAIL","PASSWORD","COMPANY","ADDRESS","CITY","ZIP_CODE","MOBILE_PHONE"});
        data.put("2", new Object[] {"someName", "SomeLastName", "SomeEmail","SomePassword","SomeCompany","SomeAddress","SomeCity","SomeZipCode","SomeMobilePhone"});
        //Iterate over data and write to sheet
        Set<String> keyset = data.keySet();
        int rownum = 0;
        for (String key : keyset)
        {
            Row row = sheet.createRow(rownum++);
            Object [] objArr = data.get(key);
            int cellnum = 0;
            for (Object obj : objArr)
            {
               Cell cell = row.createCell(cellnum++);
               if(obj instanceof String)
                    cell.setCellValue((String)obj);
                else if(obj instanceof Integer)
                    cell.setCellValue((Integer)obj);
            }
        }
        try
        {
            //Write the workbook in file system
            FileOutputStream out = new FileOutputStream(new File("ReadingWritingExcelFiles_demo.xlsx"));
            workbook.write(out);
            out.close();
            System.out.println("ReadingWritingExcelFiles_demo.xlsx written successfully on disk.");
        } 
        catch (Exception e) 
        {
            e.printStackTrace();
        }
        
    }
}
