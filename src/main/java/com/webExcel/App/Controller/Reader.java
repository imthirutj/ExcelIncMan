package com.webExcel.App.Controller;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;


@Controller
public class Reader {

	@GetMapping("/")
	public String home() {
		return "index";
	}
	
	@GetMapping("/Reader")
	public String Reader() throws IOException {
		
FileInputStream inputStream = new FileInputStream("./src/main/resources/webapp/uploads/IncidentMan.xlsx");
		
		XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
		
		
		//XSSFSheet sheet= workbook.getSheet("List of countries");
		
		//For index
		XSSFSheet sheet= workbook.getSheetAt(0);
		
		int rows=sheet.getLastRowNum();
		int cols=sheet.getRow(1).getLastCellNum();
	
		
		
		System.out.println("Rows"+rows);
		
		int arr[]= {0,1};
		int exarr[]= {8,3};
		colWriter colw=new colWriter();

		
		Object[][] Data= new Object[rows][cols];
		for(int i=0;i<arr.length;i++) {
		
		for(int r=0;r<=rows;r++) {
			XSSFRow row=sheet.getRow(r+1);
			if(row==null) {break;}
			
			for(int c=0;c<cols;c++) {
				XSSFCell cell=row.getCell(c);
				
			System.out.print("CellTYPE:"+cell.getCellType()+' ');
				switch(cell.getCellType()) {
			
				case STRING:
					String val=cell.getStringCellValue();
					 if(val.equalsIgnoreCase("3 - Moderate")) {val="P3";}
					 if(val.equalsIgnoreCase("4 - Normal")){val="P4";}
					 System.out.print(val);
					 Data[r][c]=val;
			
				
				break;
				
				case NUMERIC:
					
				 Date date =cell.getDateCellValue();
				 Data[r][c]=date;
				 System.out.print(date);break;
				case BOOLEAN:System.out.print(cell.getBooleanCellValue());break;
			
				
				default:
					break;
				}
				System.out.print(" | ");
			}
			
			
			System.out.println();
			
		
			
		}
		 colw.writeExcel(Data); 
		}workbook.close();inputStream.close();
		
		
		System.out.println("NEXXT");
		return "DownloadList";
	}
	
}
