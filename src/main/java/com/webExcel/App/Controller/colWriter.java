package com.webExcel.App.Controller;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.temporal.TemporalField;
import java.time.temporal.WeekFields;
import java.util.Arrays;
import java.util.Calendar;
import java.util.Date;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class colWriter {


	public colWriter() throws IOException {

		 XSSFWorkbook workbookCons = new XSSFWorkbook();
		
			
			 XSSFSheet sheet=workbookCons.createSheet("IncMan");
		 Object cnData[]= {"Month",	"Week"	,"Open Date",
				 "Priority"	,"Type",	"Status",	"Application Cluster",
				 "Application",	"Incident #",	"Assignment Group",	"Description",	"1-4-00 0:00"
};
		 
 
		 //create header
		 XSSFRow rowh= sheet.createRow(0);
		 int cols=cnData.length;
		 for(int ch=0;ch<cols;ch++) {
			 XSSFCell cellh= rowh.createCell(ch);
			 Object val=cnData[ch];
			 
			 CellStyle style = workbookCons.createCellStyle();
			  XSSFFont font = workbookCons.createFont();
			  font.setBold(true); style.setFont(font);
			 
			 if(val instanceof String) {
				 
				 cellh.setCellValue((String)val);
				 cellh.setCellStyle(style);
			 }
			 if(val instanceof Integer) {
				 cellh.setCellValue((Integer)val);
			 }
			 if(val instanceof Boolean) {
				 cellh.setCellValue((Boolean)val);
			 }
		 }
		FileOutputStream fileOutputStream= new FileOutputStream(".\\src\\main\\resources\\webapp\\uploads\\IncMan.xlsx");
		 workbookCons.write(fileOutputStream);
		 workbookCons.close();
		 fileOutputStream.close();
		
		
	}
	




	public void writeExcel(Object Data[][]) throws IOException {
		

		FileInputStream fis= new FileInputStream(new File(".\\src\\main\\resources\\webapp\\uploads\\IncMan.xlsx"));
		FileOutputStream fileOutputStream= new FileOutputStream(".\\src\\main\\resources\\webapp\\uploads\\IncMan.xlsx");
		

        XSSFWorkbook workbook = new XSSFWorkbook();
		
		XSSFSheet sheet=workbook.createSheet("IncMan");

	 Object cnData[]= {"Month",	"Week"	,"Open Date",
			 "Priority"	,"Type",	"Status",	"Application Cluster",
			 "Application",	"Incident #",	"Assignment Group",	"Description",	"1-4-00 0:00","Assigned To",
			 "caller","configuration Item","resolved",
			 "MTTR","Incident","Assignmet Group", "priority","state","description","opened on","closed notes","resolved",
			 "CF", "week","status","incident","Application Cluster","Application",	
};
	 

	 //create header
	 XSSFRow rowh= sheet.createRow(0);
	 int cols=cnData.length;
	 
	 
	 for(int ch=0;ch<cols;ch++) {
		 XSSFCell cellh= rowh.createCell(ch);
		 Object val=cnData[ch];
		 
		 CellStyle style = workbook.createCellStyle();
		  XSSFFont font = workbook.createFont();
		  font.setBold(true); style.setFont(font);
		  
		  sheet.autoSizeColumn(ch);
		 
		 if(val instanceof String) {
			 
			 cellh.setCellValue((String)val);
			 cellh.setCellStyle(style);
		 }
		 if(val instanceof Integer) {
			 cellh.setCellValue((Integer)val);
		 }
		 if(val instanceof Boolean) {
			 cellh.setCellValue((Boolean)val);
		 }
	 }
		
	 
	 //create body excel 
	 int arr[]= {8,3,12,5,9,10,2,13,14,15,11};
		 int rows=Data.length;
		 cols=Data[0].length;
		 
		 System.out.println("No f Rows:"+rows);
		 System.out.println("No of cols"+cols);
		 
		 for(int r=0;r<rows;r++) {
			 XSSFRow row= (XSSFRow) sheet.createRow(r+1);
			 for(int c=0;c<cols;c++) {
				 XSSFCell cell= row.createCell(arr[c]);
				 Object val=Data[r][c];
				 
				 if(c!=10) {
				 sheet.autoSizeColumn(c);}
				 if(c==10) {
				 sheet.setColumnWidth(10, 10000);}
				 
				 if(val instanceof String) {
					 cell.setCellValue((String)val);
				 }
				 if(val instanceof Integer) {
					 cell.setCellValue((Integer)val);
				 }
				 if(val instanceof Date) {
					 CreationHelper creationHelper = workbook.getCreationHelper();
					 CellStyle cellStyle = workbook.createCellStyle();
					    cellStyle.setDataFormat(creationHelper.createDataFormat().getFormat(
								"dd-mmm-yy"));
					    cell.setCellStyle(cellStyle);
					 cell.setCellValue((Date)val);
					 
				
					 
					 
				 }
				 if(val instanceof Boolean) {
					 cell.setCellValue((Boolean)val);
				 }	 
			 }
			 
			 //extra month and week
			 
			 for(int cc=0;cc<2;cc++) {
			 XSSFCell cell= row.createCell(cc);
			 Object val=Data[r][6];
			 sheet.autoSizeColumn(0);
			 
			 if(val instanceof String) {
				 cell.setCellValue((String)val);
			 }
			 if(val instanceof Integer) {
				 cell.setCellValue((Integer)val);
			 }
			 if(val instanceof Date) {
				 CreationHelper creationHelper = workbook.getCreationHelper();
				 CellStyle cellStyle = workbook.createCellStyle();
				    cellStyle.setDataFormat(creationHelper.createDataFormat().getFormat(
							"dd-mmm-yy"));
				    cell.setCellStyle(cellStyle);
				 
				 
				 DateFormat df = new SimpleDateFormat("MMM''yy");
		          Date date = (Date) val;
		          
		          if(cc==0) {
		          cell.setCellValue(df.format(date));
		          }
		          else {
		        	  DateFormat dyear = new SimpleDateFormat("yyyy");
		        	  
						 int year=Integer.parseInt(dyear.format(date));
						 
						 DateFormat dmon = new SimpleDateFormat("MM");
			        	  
						 int mon=Integer.parseInt(dmon.format(date));
						 
						 DateFormat ddate = new SimpleDateFormat("dd");
			        	  
						 int dates=Integer.parseInt(ddate.format(date));
				          // create WeekFields
		              WeekFields weekFields
		              = WeekFields.of(DayOfWeek.MONDAY, 1);
		    
		          // apply weekOfMonth()
		          TemporalField weekOfMonth
		              = weekFields.weekOfMonth();

		          LocalDate day= LocalDate.of(year,mon,dates);
		         // LocalDate day= LocalDate.of(2022,06,01);
		 
		          // get week of month for localdate
		          int wom = day.get(weekOfMonth);
		    
		          // print results
		      //    System.out.println("week of month for "+ day + " :" + wom);
		          cell.setCellValue("WK"+wom);
				        
		          }
		         
		        	  
			 }
			 if(val instanceof Boolean) {
				 cell.setCellValue((Boolean)val);
			 }	 
		 }
			 
			//desc as type 
			 
			 {
				 XSSFCell cell= row.createCell(4);
				 //desc val
				 Object val=Data[r][5];
				 
				 //status val
				 Object valSt=Data[r][3];
				 sheet.autoSizeColumn(4);
				 
				 if(val instanceof String) {
					 
					String mystr= val.toString();
					String mystatus= valSt.toString();
					
					String[] Pend={"Pending Customer","Pending Vendor","Pending Validation","Pending Change"};
					
					if(Arrays.asList(Pend).contains(mystatus)) 
					{
						 cell.setCellValue("user reported");
					}
					
					
					 if(!Arrays.asList(Pend).contains(mystatus) && (mystr.contains("UC4/2400") || mystr.contains("SPLUNK")))
					 {
					 cell.setCellValue("automated");
					 }
					 
					 if(!Arrays.asList(Pend).contains(mystatus) && !(mystr.contains("UC4/2400") || mystr.contains("SPLUNK"))) {
						 cell.setCellValue("user reported");
					 }
					 
				 }
			 }
			 
			 
			 //Application
			 {
				 XSSFCell cell= row.createCell(7);
				 XSSFCell cell_clus= row.createCell(6);
				 //desc val
				 Object valConfig=Data[r][8];
				 Object valAss=Data[r][4];
				 
		
				 sheet.autoSizeColumn(7);
				 
				 if(valConfig instanceof String) {
					 
					 
					String myCon= valConfig.toString();
					String myAss= valAss.toString();
					//for wms wcs ass-grp
					if(myAss.equalsIgnoreCase("app-global-wmssup") ||
							myAss.equalsIgnoreCase("app-flwdw-wcssup"))
					{
						cell_clus.setCellValue("WM O&F");
						if(myCon.toUpperCase().contains("WDW") || myCon.toUpperCase().contains("WCS") )
						{
							cell.setCellValue("WDW");
						}
						
						else if(myCon.toUpperCase().contains("DLR"))
						{
							cell.setCellValue("DLR");
						}

						else
						{
							cell.setCellValue("DLR");
						}
					}
					//for sdp ass-grp
					else if(myAss.equalsIgnoreCase("app-flwdw-sdp")){
						cell_clus.setCellValue("Merchandise Operation");
						cell.setCellValue("SDP");
					}
					//for doc direct ass-grp
					else if(myAss.equalsIgnoreCase("app-global-DocDirec")){
						cell_clus.setCellValue("WM O&F");
						cell.setCellValue("DOC DIRECT");
					}
					//for orbatch,oretail,orrib ass-grp
					else if(myAss.equalsIgnoreCase("app-global-orbatch")
							|| myAss.equalsIgnoreCase("app-global-oretail")
							|| myAss.equalsIgnoreCase("app-global-orrib")){
						cell_clus.setCellValue("Merchandise Operation");
						cell.setCellValue("RMS");
					}
					//for pktrack ass-grp
					else if(myAss.equalsIgnoreCase("app-global-pktrack")){
						cell_clus.setCellValue("WM O&F");
						cell.setCellValue("WHSSYS");
					}
					//for pride ass-grp
					else if(myAss.equalsIgnoreCase("app-global-pride")){
						cell_clus.setCellValue("Merchandise Operation");
						cell.setCellValue("Pride");
					}
					//for rpas ass-grp
					else if(myAss.equalsIgnoreCase("app-global-rpas")){
						cell_clus.setCellValue("Merchandise Operation");
						cell.setCellValue("RPAS");
					}
					//for sim ass-grp
					else if(myAss.equalsIgnoreCase("app-global-sim")){
						cell_clus.setCellValue("Merchandise Operation");
						cell.setCellValue("SIM");
					}
					//for whshw ass-grp
					else if(myAss.equalsIgnoreCase("app-global-whshw")){
						cell_clus.setCellValue("WM O&F");
						cell.setCellValue("WHSHW");
					}
					//for whssys ass-grp
					else if(myAss.equalsIgnoreCase("app-global-whssyss")){
						cell_clus.setCellValue("WM O&F");
						cell.setCellValue("WHSSYS");
					}
				 }
			 }
			 
			
		//for mttr
			 int arMt[]= {17,18,19,20,21,22,23,24};
			 int rowML=Data.length;
			 cols=Data[0].length;
			 int arMc[]= {0,4,1,3,5,6,10,9};
			// System.out.println(Data[r][9]);
			
				
				 for(int c=0;c<arMt.length;c++) {
					 XSSFCell cell= row.createCell(arMt[c]);
					// System.out.println("arMc: "+c+" "+arMc[c]);
					 Object val=Data[r][arMc[c]];
					 
					 if(c!=10) {
					 sheet.autoSizeColumn(c);}
					 if(c==10) {
					 sheet.setColumnWidth(10, 10000);}
					 
					 if(val instanceof String) {
						 cell.setCellValue((String)val);
					 }
					 if(val instanceof Integer) {
						 cell.setCellValue((Integer)val);
					 }
					 if(val instanceof Date) {
						 CreationHelper creationHelper = workbook.getCreationHelper();
						 CellStyle cellStyle = workbook.createCellStyle();
						    cellStyle.setDataFormat(creationHelper.createDataFormat().getFormat(
									"dd-mmm-yy"));
						    cell.setCellStyle(cellStyle);
						 cell.setCellValue((Date)val);
						 
					
						 
						 
					 }
					 if(val instanceof Boolean) {
						 cell.setCellValue((Boolean)val);
					 }	 
				 }
				 
				 //FOR CF
				 //0-Incident 3-state 6-opened on 
				 {
				 int[] arCF= {28,27,26};
				 int rowCF=Data.length;
				 cols=Data[0].length;
				 int arCFs[]= {0,3,6};
				 

				 for(int c=0;c<arCF.length;c++) {
					 XSSFCell cell= row.createCell(arCF[c]);
					// System.out.println("arMc: "+c+" "+arMc[c]);
					 Object val=Data[r][arCFs[c]];
					 
					 if(c!=10) {
					 sheet.autoSizeColumn(c);}
					 if(c==10) {
					 sheet.setColumnWidth(10, 10000);}
					 
					 if(val instanceof String) {
						 cell.setCellValue((String)val);
					 }
					 if(val instanceof Integer) {
						 cell.setCellValue((Integer)val);
					 }
					 if(val instanceof Date) {
						 CreationHelper creationHelper = workbook.getCreationHelper();
						 CellStyle cellStyle = workbook.createCellStyle();
						    cellStyle.setDataFormat(creationHelper.createDataFormat().getFormat(
									"dd-mmm-yy"));
						    cell.setCellStyle(cellStyle);
						 
						 
						 DateFormat df = new SimpleDateFormat("MMM''yy");
				          Date date = (Date) val;
				          
				          if(c==0) {
				          cell.setCellValue(df.format(date));
				          }
				          else {
				        	  DateFormat dyear = new SimpleDateFormat("yyyy");
				        	  
								 int year=Integer.parseInt(dyear.format(date));
								 
								 DateFormat dmon = new SimpleDateFormat("MM");
					        	  
								 int mon=Integer.parseInt(dmon.format(date));
								 
								 DateFormat ddate = new SimpleDateFormat("dd");
					        	  
								 int dates=Integer.parseInt(ddate.format(date));
						          // create WeekFields
				              WeekFields weekFields
				              = WeekFields.of(DayOfWeek.MONDAY, 1);
				    
				          // apply weekOfMonth()
				          TemporalField weekOfMonth
				              = weekFields.weekOfMonth();

				          LocalDate day= LocalDate.of(year,mon,dates);
				         // LocalDate day= LocalDate.of(2022,06,01);
				 
				          // get week of month for localdate
				          int wom = day.get(weekOfMonth);
				    
				          // print results
				      //    System.out.println("week of month for "+ day + " :" + wom);
				          cell.setCellValue("WK"+wom);
						        
				          }
				         
				        	  
					 }
					 if(val instanceof Boolean) {
						 cell.setCellValue((Boolean)val);
					 }	 
				 }
				 
				 }
				 //Application and clus for cf
				 {
					 XSSFCell cell= row.createCell(30);
					 XSSFCell cell_clus= row.createCell(29);
					 //desc val
					 Object valConfig=Data[r][8];
					 Object valAss=Data[r][4];
					 
			
					 sheet.autoSizeColumn(7);
					 
					 if(valConfig instanceof String) {
						 
						 
						String myCon= valConfig.toString();
						String myAss= valAss.toString();
						//for wms wcs ass-grp
						if(myAss.equalsIgnoreCase("app-global-wmssup") ||
								myAss.equalsIgnoreCase("app-flwdw-wcssup"))
						{
							cell_clus.setCellValue("WM O&F");
							if(myCon.toUpperCase().contains("WDW") || myCon.toUpperCase().contains("WCS") )
							{
								cell.setCellValue("WDW");
							}
							
							else if(myCon.toUpperCase().contains("DLR"))
							{
								cell.setCellValue("DLR");
							}

							else
							{
								cell.setCellValue("DLR");
							}
						}
						//for sdp ass-grp
						else if(myAss.equalsIgnoreCase("app-flwdw-sdp")){
							cell_clus.setCellValue("Merchandise Operation");
							cell.setCellValue("SDP");
						}
						//for doc direct ass-grp
						else if(myAss.equalsIgnoreCase("app-global-DocDirec")){
							cell_clus.setCellValue("WM O&F");
							cell.setCellValue("DOC DIRECT");
						}
						//for orbatch,oretail,orrib ass-grp
						else if(myAss.equalsIgnoreCase("app-global-orbatch")
								|| myAss.equalsIgnoreCase("app-global-oretail")
								|| myAss.equalsIgnoreCase("app-global-orrib")){
							cell_clus.setCellValue("Merchandise Operation");
							cell.setCellValue("RMS");
						}
						//for pktrack ass-grp
						else if(myAss.equalsIgnoreCase("app-global-pktrack")){
							cell_clus.setCellValue("WM O&F");
							cell.setCellValue("WHSSYS");
						}
						//for pride ass-grp
						else if(myAss.equalsIgnoreCase("app-global-pride")){
							cell_clus.setCellValue("Merchandise Operation");
							cell.setCellValue("Pride");
						}
						//for rpas ass-grp
						else if(myAss.equalsIgnoreCase("app-global-rpas")){
							cell_clus.setCellValue("Merchandise Operation");
							cell.setCellValue("RPAS");
						}
						//for sim ass-grp
						else if(myAss.equalsIgnoreCase("app-global-sim")){
							cell_clus.setCellValue("Merchandise Operation");
							cell.setCellValue("SIM");
						}
						//for whshw ass-grp
						else if(myAss.equalsIgnoreCase("app-global-whshw")){
							cell_clus.setCellValue("WM O&F");
							cell.setCellValue("WHSHW");
						}
						//for whssys ass-grp
						else if(myAss.equalsIgnoreCase("app-global-whssyss")){
							cell_clus.setCellValue("WM O&F");
							cell.setCellValue("WHSSYS");
						}
					 }
				 }
				 
				 
			 
			
		
		 
		 
		 }
		 fis.close();
		 workbook.write(fileOutputStream);
		 workbook.close();
		 fileOutputStream.close();
		 System.out.println("Successs");
		 
		
		 
		 
		 
	
	}
	
	
	

	
	public int getWeekNum(String input) throws ParseException {
		  String sDate1=input;  
		    Date date1=new SimpleDateFormat("dd-MM-yyyy").parse(sDate1);  
	      Calendar cl = Calendar. getInstance();
	      cl.setTime(date1);
	      
	      int d=cl.WEEK_OF_MONTH;
	      System.out.println("today is a "+cl.WEEK_OF_MONTH +"week of the month");
	    return d;
	}

}
