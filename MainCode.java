package com.src.apachePackage;

import java.io.*; 
import org.apache.poi.hssf.usermodel.HSSFWorkbook; 
import org.apache.poi.ss.usermodel.Sheet; 
import org.apache.poi.ss.usermodel.Workbook; 

public class MainCode {
	public static void main(String[] args)  
		      throws FileNotFoundException, IOException 
		    { 
		  
		        // Creating Workbook instances 
		        Workbook wb = new HSSFWorkbook(); 
		  
		        // An output stream accepts output bytes and sends them to sink. 
		        OutputStream fileOut = new FileOutputStream("Geeks.xlsx"); 
		          
		        // Creating Sheets using sheet object 
		        Sheet sheet1 = wb.createSheet("Array"); 
		        Sheet sheet2 = wb.createSheet("String"); 
		        Sheet sheet3 = wb.createSheet("LinkedList"); 
		        Sheet sheet4 = wb.createSheet("Tree"); 
		        Sheet sheet5 = wb.createSheet("Dynamic Programing"); 
		        Sheet sheet6 = wb.createSheet("Puzzles"); 
		          
		  
		System.out.println("Sheets Has been Created successfully"); 
		  
		        wb.write(fileOut); 
		    } 
}
