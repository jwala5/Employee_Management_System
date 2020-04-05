package dao;
import model.*;
import java.util.*;
import java.io.*;

import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excelfile {
	static Workbook wb;
	static Sheet sh;
	static Row row;
	static Cell cell;
	static FileInputStream in;
	static FileOutputStream out;
	   int i=0;
	
	   
	   //Insert instrument details in excel file	
	public void excelgenerator(Employee employee, List<Employee> list) throws Exception {
		
		try {
			in = new FileInputStream("EmpSheet.xlsx");
			XSSFWorkbook workbook = new XSSFWorkbook(in);
			XSSFSheet sheet = workbook.getSheetAt(0);
			int rowCount = sheet.getLastRowNum();
			 
			for(Employee employee1 :list) {
				 Row row = sheet.createRow(++rowCount);
				 int columnCount = 0;
				 Cell cell = row.createCell(columnCount);
				 cell.setCellValue(rowCount);
				 
				 row.createCell(0).setCellValue(employee1.getId());
				 row.createCell(1).setCellValue(employee1.getEmpname());
				 row.createCell(2).setCellValue(employee1.getDept()); 
				 row.createCell(3).setCellValue(employee1.getSalary());
				 row.createCell(4).setCellValue(employee1.getAge());
				 row.createCell(5).setCellValue(employee1.getPost()); 
			 	 }
			   
		    out = new FileOutputStream(new File("EmpSheet.xlsx"));
	        workbook.write(out);
	        out.close();
	                    
		    }catch (Exception e) {
		    	System.out.println(e.getMessage());
		    }
	
	}
//Read All Data in Excel file
	public void excelreader(String fname) {
		try
      {
          FileInputStream file = new FileInputStream(new File(fname));
          XSSFWorkbook workbook = new XSSFWorkbook(file);
          XSSFSheet sheet = workbook.getSheetAt(0);
          Iterator<Row> rowIterator = sheet.iterator();
        int i=0;
          while (rowIterator.hasNext()) 
          {
              row = rowIterator.next();
              if(i!=1) {
              	 row = rowIterator.next();
              	 i++;
              }
              //For each row, iterate through all the columns
              Iterator<Cell> cellIterator = row.cellIterator();
               
              while (cellIterator.hasNext()) 
              {
                  Cell cell = cellIterator.next();
                  //Check the cell type and format accordingly
                  switch (cell.getCellType()) 
                  {
                      //case Cell.CELL_TYPE_NUMERIC:
                         // System.out.print(cell.getNumericCellValue() + "\t");
                          //break;
                      case Cell.CELL_TYPE_STRING:
                          System.out.print(cell.getStringCellValue() + "\t\t");
                          break;
                  }
              }
              System.out.println("");
          }
          file.close();
      
      }catch (Exception e) {
      	System.out.println(e.getMessage());
      }
		
	}
	//	

//	//delete item from excel file
	public void deleteitem(String id) {
		try
	      {
			
	          FileInputStream file = new FileInputStream(new File("EmpSheet.xlsx"));
	          XSSFWorkbook workbook = new XSSFWorkbook(file);
	          XSSFSheet sheet = workbook.getSheetAt(0);
	          for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
	        	  row = sheet.getRow(rowIndex);
	        	  if (row != null) {
	        	    Cell cell = row.getCell(0);
	        	    String value=cell.getStringCellValue();
	        	    if (value.equals(id)) {
	        	    	sheet.removeRow(sheet.getRow(rowIndex));
	        	    	//sheet.shiftRows(rowIndex+1,sheet.getLastRowNum(), -1);
	        	    	 System.out.println("Deleted Successfully!");
	        	    }
	        	  }
	        	}
	            out = new FileOutputStream(new File("EmpSheet.xlsx"));
		        workbook.write(out);
		        out.close();
	     
	       }catch (Exception e) {
	        	System.out.println(e.getMessage());
	      }
	}



}