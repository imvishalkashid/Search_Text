package com.test.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Iterator;
import java.util.List;
import java.util.Set;
import java.util.concurrent.TimeUnit;
import java.util.logging.Level;
import java.util.logging.Logger;

import org.apache.commons.logging.Log;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.firefox.FirefoxOptions;
import org.openqa.selenium.logging.LogType;
import org.openqa.selenium.logging.LoggingPreferences;
import org.openqa.selenium.remote.CapabilityType;
import org.openqa.selenium.remote.DesiredCapabilities;

import com.google.common.io.Files;
import com.sun.jna.platform.FileUtils;

public class ReadExcel {
public static WebDriver driver;


	
	private static Logger Log = Logger.getLogger(Log.class.getName());
	
	 public void readExcel(String filePath,String fileName,String sheetName) throws IOException{
		
		 LoggingPreferences pref = new LoggingPreferences();
		    pref.enable(LogType.BROWSER, Level.OFF);
		    pref.enable(LogType.CLIENT, Level.OFF);
		    pref.enable(LogType.DRIVER, Level.OFF);
		    pref.enable(LogType.PERFORMANCE, Level.OFF);
		    pref.enable(LogType.PROFILER, Level.OFF);
		    pref.enable(LogType.SERVER, Level.OFF);


		    DesiredCapabilities desiredCapabilities = DesiredCapabilities.firefox();
		    
		    desiredCapabilities.setCapability(CapabilityType.LOGGING_PREFS, pref);
		 
		    System.setProperty("webdriver.gecko.driver","C:/Vishal/geckodriver.exe");
			driver = new FirefoxDriver(desiredCapabilities);
			Log.info("New driver instantiated");
			driver.manage().timeouts().implicitlyWait(1, TimeUnit.MINUTES);
		 
		 
		 
			//Create an object of File class to open xlsx file
			File file =    new File(filePath+"\\"+fileName);	 
		    
			//Create an object of FileInputStream class to read excel file
			FileInputStream inputStream = new FileInputStream(file);    
		    
		    
		    Workbook guru99Workbook = null;
		    
		    //Find the file extension by splitting file name in substring  and getting only extension name
		    String fileExtensionName = fileName.substring(fileName.indexOf("."));
		    
		    //Check condition if the file is xlsx file
		    if(fileExtensionName.equals(".xlsx")){

		    //If it is xlsx file then create object of XSSFWorkbook class
		    guru99Workbook = new XSSFWorkbook(inputStream);

		    }

		    //Check condition if the file is xls file
		    else if(fileExtensionName.equals(".xls")){

		    //If it is xls file then create object of XSSFWorkbook class
		    guru99Workbook = new HSSFWorkbook(inputStream);
		    		
		    }
		    
		    //Read sheet inside the workbook by its name
		    Sheet guru99Sheet = guru99Workbook.getSheet(sheetName);

		    //Find number of rows in excel file
		    int rowCount = guru99Sheet.getLastRowNum()-guru99Sheet.getFirstRowNum();
		    
		    //Create a loop over all the rows of excel file to read it

		    List<String> list1 = new ArrayList<String>();
			List<String> list2 = new ArrayList<String>();
			
			//System.out.println(list1);
		    
		    for (int i = 0; i < rowCount+1; i++) {
		    	
		    	guru99Sheet.getRow(i);

		        Row row = guru99Sheet.getRow(i);
		        
		        
		        
		        
		      //Create a loop to print cell values in a row

		        for (int j = 0; j < row.getLastCellNum(); j++) {
		        	
		        	
		        	
		            //Print Excel data in console
		        	
		        	
		        	String sCell_value=row.getCell(j).getRichStringCellValue().toString();
		            //System.out.print(sCell_value+"|");
		            
		          
		        	//((JavascriptExecutor)driver).executeScript("window.open('" + sCell_value + " ')");
		            
		            driver.navigate().to(sCell_value);
		       
		                 
		            if((driver.getPageSource().contains("Rajesh Ghonasgi"))&&(driver.getPageSource().contains("Rajesh Ghonasgi is the Chief Financial Officer")))
		            	{
			       	 	
		            			System.out.println(driver.getCurrentUrl());
			       	  			list1.add(driver.getCurrentUrl()); 
			       	  			
			       	  			
			       	  		    try {
										Thread.sleep(3000);
									} catch (InterruptedException e) 
									{
										// TODO Auto-generated catch block
										e.printStackTrace();
									}
			       	  
			         	}	
		            
		            
		            else
	         			{
	       	 
	         				System.out.println(driver.getCurrentUrl());
	         				list2.add(driver.getCurrentUrl()); 
	         			}
			    }
		       

		            //System.out.print(row.getCell(j).getStringCellValue()+"|| ");

		        } 
		    
		    
	        System.out.println("+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++");
		     System.out.println("Below Link where Text is present");
		     
		     int listzise= list1.size();
  	  		 //System.out.println(listzise);
		     
		    
	    	  if(listzise==0)
	    	  {
	    		  System.out.println("*************************************");
	    		  System.out.println("*************************************");
	    		  
	    		  System.out.println("There is no any link wrt.Search text");
	    	   
	              System.out.println("*************************************");
	              System.out.println("*************************************");
	    	  } else
	    	  {
		      for (String element : list1) 
		       {
		    	  
		    	  
		    	  System.out.println(element);
		    	  	  
		      	}
	    	  } 
		    
		      System.out.println("+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++");
		      
		      System.out.println("Below links where Text is not present ");
		      for (String element : list2) 
		      {
		    	    System.out.println(element);
		      }
	            
		       // System.out.println();

	 }
	 

	 public static void main(String args[]) throws IOException{

	    //Create an object of ReadGuru99ExcelFile class
		 ReadExcel objExcelFile = new ReadExcel();

	    //Prepare the path of excel file

	    String filePath = "E:/";
	    
	    String sheetName ="Seqriteone";
	    
	    String fileName = "Study Material/test/readfile.xls";

	    //Call read file method of the class to read data

	    objExcelFile.readExcel(filePath,fileName,sheetName);

	    }
 
	 
	 
}
