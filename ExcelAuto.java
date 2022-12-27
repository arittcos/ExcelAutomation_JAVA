package com.excelauto;

import java.io.*;
import java.util.*;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFCell;
 

public class ExcelAuto {

	public static void main(String[] args)throws FileNotFoundException, IOException {
		// TODO Auto-generated method stub
		
		XSSFWorkbook wb = new XSSFWorkbook();
		
		//adding two sheets to the workbook
		XSSFSheet sheet1= wb.createSheet("Student_Details_1");
		XSSFSheet sheet2= wb.createSheet("Student_Details_2");
		
		//adding data in sheet2
		Map<String, Object[]> data = new TreeMap<String, Object[]>();
        data.put("1", new Object[] {"ID", "NAME", "LASTNAME"});
        data.put("2", new Object[] {1, "Amit", "Shukla"});
        data.put("3", new Object[] {2, "Lokesh", "Gupta"});
        data.put("4", new Object[] {3, "John", "Adwards"});
        data.put("5", new Object[] {4, "Brian", "Schultz"});
          
        //Iterate over data and write to sheet
        Set<String> keyset = data.keySet();
        int rownum = 0;
        for (String key : keyset)
        {
            XSSFRow row = sheet1.createRow(rownum++);
            Object [] objArr = data.get(key);
            int cellnum = 0;
            for (Object obj : objArr)
            {
               XSSFCell cell = row.createCell(cellnum++);
               if(obj instanceof String)
                    cell.setCellValue((String)obj);
                else if(obj instanceof Integer)
                    cell.setCellValue((Integer)obj);
            }
        }
        
        //adding data in sheet2
        Map<String, Object[]> data1 = new TreeMap<String, Object[]>();
        data1.put("1", new Object[] {"ID", "NAME", "LASTNAME"});
        data1.put("2", new Object[] {1, "Amit", "Shukla"});
        data1.put("3", new Object[] {2, "Lokesh", "Gupta"});
        data1.put("4", new Object[] {3, "John", "Adwards"});
        data1.put("5", new Object[] {4, "Brian", "Schultz"});
        data1.put("6", new Object[] {5, "Amit", "Shukla"});
        data1.put("7", new Object[] {6, "Lokesh", "Gupta"});
        data1.put("8", new Object[] {7, "John", "Adwards"});
        data1.put("9", new Object[] {8, "Brian", "Schultz"});
          
        //Iterate over data and write to sheet
        Set<String> keyset1 = data1.keySet();
        int rownum1 = 0;
        for (String key : keyset1)
        {
            XSSFRow row = sheet2.createRow(rownum1++);
            Object [] objArr = data1.get(key);
            int cellnum = 0;
            for (Object obj : objArr)
            {
               XSSFCell cell = row.createCell(cellnum++);
               if(obj instanceof String)
                    cell.setCellValue((String)obj);
                else if(obj instanceof Integer)
                    cell.setCellValue((Integer)obj);
            }
        }
		
		try
		{
			//to create object of scanner class
			Scanner myObj = new Scanner(System.in);
			
			//to read the fileName from user
		    System.out.println("Enter filename"); 
		    String fileName = myObj.nextLine();
		    
		    //creating new excel file with user defined name
			String filename = "E:\\javaExcel\\"+fileName+".xlsx";
			FileOutputStream fileOut = new FileOutputStream(filename);
			
			wb.write(fileOut);
			
			fileOut.close();
			
			
		}
		catch(Exception e)
		{
			System.out.println(e);
		}
		
	}

}
