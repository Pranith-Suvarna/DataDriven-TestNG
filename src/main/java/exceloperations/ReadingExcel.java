package exceloperations;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.xssf.usermodel.*;

public class ReadingExcel {

	public static void main(String[] args) throws IOException {
		
		String excelFilePath=".\\datafiles\\countries.xlsx";             //Create a string variable and store the excel path   
		FileInputStream inputstream=new FileInputStream(excelFilePath);  //When we open the file in reading mode we use FileInputStream class which will create a stream into this class and read the data
		
		XSSFWorkbook workbook=new XSSFWorkbook(inputstream);             //Then we use XSSFWorkbook class which will get the workbook from the excel file
		XSSFSheet sheet=workbook.getSheet("Sheet1");                     //Now from the above workbook we have to get the sheet through sheet name
		//XSSFSheet sheet=workbook.getSheetAt(0);	                     //Now from the above workbook we have to get the sheet through sheet index
		
		////  USING FOR LOOP
		
	/*	int rows=sheet.getLastRowNum();                                //to get the no. of rows
		int cols=sheet.getRow(1).getLastCellNum();                     //to get the no. of columns
		
		for(int r=0;r<=rows;r++)                                       //outer loop representing the rows
		{
			XSSFRow row=sheet.getRow(r); //0
			
			for(int c=0;c<cols;c++)                                    //inner loop representing the cells in a row
			{
				XSSFCell cell=row.getCell(c);
				
				switch(cell.getCellType())                             // to check whether the cell type is string , numeric or boolean
				{
				case STRING: System.out.print(cell.getStringCellValue()); break;
				case NUMERIC: System.out.print(cell.getNumericCellValue());break;
				case BOOLEAN: System.out.print(cell.getBooleanCellValue()); break;
				}
				System.out.print(" | ");
			}
			System.out.println();
		}
    */
		
		
		///////// Iterator ////////////////////////
		
		Iterator iterator=sheet.iterator();
		
		while(iterator.hasNext())                                                // check if object is there or not in the iterator
		{
			XSSFRow row=(XSSFRow) iterator.next();                               //will return the rows in the sheet
			
			Iterator cellIterator=row.cellIterator();                            // capture all the cells in the row
			
			while(cellIterator.hasNext())                                        //will check if cell is there or not int the row
			{
				XSSFCell cell=(XSSFCell) cellIterator.next();
				
				switch(cell.getCellType())                                       // to check whether the cell type is string , numeric or boolean
				{
				case STRING: System.out.print(cell.getStringCellValue()); break;
				case NUMERIC: System.out.print(cell.getNumericCellValue());break;
				case BOOLEAN: System.out.print(cell.getBooleanCellValue()); break;
				}
				System.out.print(" |  ");
			}
			System.out.println();
		}
				
		
	}

}
