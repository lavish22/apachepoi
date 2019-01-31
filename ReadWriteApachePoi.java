package com.src.apachePackage;
import java.io.*;
import java.util.*;
import java.sql.*; 
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadWriteApachePoi {    
	public static void main(String[] args) 
    { 
        try { 
            FileInputStream file = new FileInputStream(new File("Geeks.xlsx")); 
  
            // Create Workbook instance holding reference to .xlsx file 
            XSSFWorkbook workbook = new XSSFWorkbook(file); 
  
            // Get first/desired sheet from the workbook 
            XSSFSheet sheet = workbook.getSheetAt(0); 
  
            // Iterate through each rows one by one
            
            List<Integer> passlist = new ArrayList<Integer>();
            List<Integer> faillist = new ArrayList<Integer>();
            
            Iterator<Row> rowIterator = sheet.iterator(); 
            while (rowIterator.hasNext()) { 
                Row row = rowIterator.next(); 
                // For each row, iterate through all the columns 
                Iterator<Cell> cellIterator = row.cellIterator(); 
                int temp=0;
                while (cellIterator.hasNext()) { 
                    Cell cell = cellIterator.next(); 
                    // Check the cell type and format accordingly 
                    switch (cell.getCellType()) { 
                    case NUMERIC: 
                        temp=(int)cell.getNumericCellValue(); 
                        break; 
                    case STRING: 
                        System.out.print(cell.getStringCellValue() + "\t"); 
                        if( cell.getStringCellValue().equals("pass") )
                        	passlist.add(temp);
                        else if( cell.getStringCellValue().equals("fail") )	
                        	faillist.add(temp);
                        break; 
                    } 
                } 
                System.out.println(""); 
            } 
            file.close();
            System.out.print(passlist);
            System.out.print(faillist);
            
            XSSFWorkbook workbookw = new XSSFWorkbook(); 
            
            // Create a blank sheet 
            XSSFSheet sheetw = workbookw.createSheet("sorted details");
            int rownum=0;
            for( Integer i : passlist ) {
            	int cellnum=0;
            	Row row = sheetw.createRow(rownum++); 
            	Cell cell1 = row.createCell(cellnum++); 
            	cell1.setCellValue(i);
            	Cell cell2 = row.createCell(cellnum++);
            	cell2.setCellValue((String)"pass");
            }
            for( Integer i : faillist ) {
            	int cellnum=0;
            	Row row = sheetw.createRow(rownum++); 
            	row.createCell(cellnum++).setCellValue(i);
            	row.createCell(cellnum++).setCellValue((String)"fail");;
            }
            try { 
                // this Writes the workbook gfgcontribute 
                FileOutputStream out = new FileOutputStream(new File("Test_results.xlsx")); 
                workbookw.write(out); 
                out.close(); 
                System.out.println("Test_results.xlsx written successfully on disk."); 
            } 
            catch (Exception e) { 
                e.printStackTrace(); 
            }
           
        } 
        catch (Exception e) { 
            e.printStackTrace(); 
        } 
        
    } 
     
}