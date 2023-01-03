package com.excelAuto;

import java.io.*;
import java.util.*;

//import org.apache.poi.xssf.usermodel.XSSFSheet;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//import org.apache.poi.xssf.usermodel.XSSFRow;
//import org.apache.poi.xssf.usermodel.XSSFCell;


public class ExcelAuto {

	public static void main(String[] args)throws FileNotFoundException, IOException {
		// TODO Auto-generated method stub
		
		FileReader reader =new FileReader("config.properties");
		Properties p = new Properties();
		p.load(reader);
		
		System.out.println(p.getProperty("user"));  
	    System.out.println(p.getProperty("password"));  
		
//		XSSFWorkbook wb = new XSSFWorkbook();
//		
//		//adding two sheets to the workbook
//		XSSFSheet sheet1= wb.createSheet("Student_Details_1");
//		XSSFSheet sheet2= wb.createSheet("Student_Details_2");
//		
//		//adding data in sheet2
//		Map<String, Object[]> data = new TreeMap<String, Object[]>();
//        data.put("1", new Object[] {"ID", "NAME", "LASTNAME"});
//        data.put("2", new Object[] {1, "Amit", "Shukla"});
//        data.put("3", new Object[] {2, "Lokesh", "Gupta"});
//        data.put("4", new Object[] {3, "John", "Adwards"});
//        data.put("5", new Object[] {4, "Brian", "Schultz"});
//          
//        //Iterate over data and write to sheet
//        Set<String> keyset = data.keySet();
//        int rownum = 0;
//        for (String key : keyset)
//        {
//            XSSFRow row = sheet1.createRow(rownum++);
//            Object [] objArr = data.get(key);
//            int cellnum = 0;
//            for (Object obj : objArr)
//            {
//               XSSFCell cell = row.createCell(cellnum++);
//               if(obj instanceof String)
//                    cell.setCellValue((String)obj);
//                else if(obj instanceof Integer)
//                    cell.setCellValue((Integer)obj);
//            }
//        }
//        
//        //adding data in sheet2
//        Map<String, Object[]> data1 = new TreeMap<String, Object[]>();
//        data1.put("1", new Object[] {"ID", "NAME", "LASTNAME"});
//        data1.put("2", new Object[] {1, "Amit", "Shukla"});
//        data1.put("3", new Object[] {2, "Lokesh", "Gupta"});
//        data1.put("4", new Object[] {3, "John", "Adwards"});
//        data1.put("5", new Object[] {4, "Brian", "Schultz"});
//          
//        //Iterate over data and write to sheet
//        Set<String> keyset1 = data1.keySet();
//        int rownum1 = 0;
//        for (String key : keyset1)
//        {
//            XSSFRow row = sheet2.createRow(rownum1++);
//            Object [] objArr = data1.get(key);
//            int cellnum = 0;
//            for (Object obj : objArr)
//            {
//               XSSFCell cell = row.createCell(cellnum++);
//               if(obj instanceof String)
//                    cell.setCellValue((String)obj);
//                else if(obj instanceof Integer)
//                    cell.setCellValue((Integer)obj);
//            }
//        }
//		
//		try
//		{
//			//to create object of scanner class
//			Scanner myObj = new Scanner(System.in);
//			Scanner myobj1 = new Scanner(System.in);
//			
//			//to read in which drive user wants to create the file
//			System.out.println("Enter in which Drive waant save (except c drive) : "); 
//		    String fileDrive = myobj1.nextLine();
//			
//			//to read the fileName from user
//		    System.out.println("Enter filename : "); 
//		    String fileName = myObj.nextLine();
//		    
//		    //creating new excel file with user defined name
//			String filename = fileDrive+":"+"\\"+fileName+".xlsx";
//			FileOutputStream fileOut = new FileOutputStream(filename);
//			
//			wb.write(fileOut);
//			
//			fileOut.close();
//			
//			//comparing two sheets
//			if(keyset.size() == keyset1.size())
//			{
//				int noMatch=0;
//				
//				for(int i=1;i<keyset.size();i++)
//				{
//					XSSFRow sheet1_row = sheet1.getRow(i);
//					XSSFRow sheet2_row = sheet2.getRow(i);
//					
//					XSSFCell sheet1_cell = sheet1_row.getCell(1);
//					XSSFCell sheet2_cell = sheet2_row.getCell(1);
//					
//					String sheet1_cell_data = sheet1_cell.toString();
//					String sheet2_cell_data = sheet2_cell.toString();
//					
//					if(sheet1_cell_data == sheet2_cell_data)
//					{
//						noMatch=0;
////						System.out.println(sheet1_cell);
////						System.out.println(sheet2_cell);
////						System.out.println(noMatch);
//					}
//					else
//					{
//						noMatch=1;
//						break;
////						System.out.println(sheet1_cell);
////						System.out.println(sheet2_cell);
////						System.out.println(noMatch);
//					}
//				}
//				
//				if(noMatch==0)
//				{
//					System.out.println("Hurry All Values Are Similar!!");
//				}
//				else
//				{
//					System.out.println("Sorry All Values Are Not Similar!!");
//				}
//			}
//			else
//			{
//				System.out.println("Two Sheets Have Different Heights!!");
//			}
//			
//			
//		}
//		catch(Exception e)
//		{
//			System.out.println(e);
//		}
//		
	}

}