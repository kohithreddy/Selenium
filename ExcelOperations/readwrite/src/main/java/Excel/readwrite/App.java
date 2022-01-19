package Excel.readwrite;

/**
 * Hello world!
 *
 */

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class App 
{
	public static void main(String[] args) throws IOException 
	{
		// TODO Auto-generated method stub
		
		File f = new File ("data.xlsx");
		
		FileInputStream Fis = new FileInputStream (f);
		
		XSSFWorkbook excelWorkBook = new XSSFWorkbook(Fis);
		
		XSSFSheet excelSheet  = excelWorkBook.getSheetAt(0);
		
		int rows = excelSheet.getPhysicalNumberOfRows();//3
		int cols = excelSheet.getRow(0).getPhysicalNumberOfCells();//2
		String data[][]= new String[rows][cols];
		XSSFCell cell;
		for(int i =1 ; i< rows;i++)
		{
			Row row = excelSheet.getRow(i);
			
			for(int j=1;j<cols;j++)
			{
			
				int k = j+1;
				
				System.out.println("User Name is " +  row.getCell(j-1).getStringCellValue() +" \t Passwod is "+ row.getCell(j).getStringCellValue() );
			}
		}
		Fis.close();	
	}
	// ******** Reference links ******************//
	
	//https://www.youtube.com/watch?v=feNbe8T8Xck&ab_channel=total-qa
	//http://total-qa.com/read-xlsx-using-apache-poi-maven-project/
	//https://www.softwaretestingmaterial.com/read-excel-files-using-apache-poi/
	//https://www.guru99.com/all-about-excel-in-selenium-poi-jxl.html#1/
	
	// *******************************************//
	
}

