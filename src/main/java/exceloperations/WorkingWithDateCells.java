package exceloperations;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCreationHelper;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WorkingWithDateCells {

	public static void main(String[] args) throws IOException {

		
		//Create a blank workbook
		XSSFWorkbook workbook = new XSSFWorkbook();
		
		
		//Create a blank sheet
		XSSFSheet sheet = workbook.createSheet("Date Formats");
		
		XSSFCell cell = sheet.createRow(0).createCell(0);       //Create a cell
		cell.setCellValue(new Date());                          //Update the cell with current date but in number format
		
		
		XSSFCreationHelper creationhelper = workbook.getCreationHelper();

		//to update the cell with current date in proper date format
		CellStyle style=workbook.createCellStyle();
	    style.setDataFormat(creationhelper.createDataFormat().getFormat("dd-mm-yyyy"));  //specify the data format
	    
	    
	    XSSFCell cell1 = sheet.createRow(1).createCell(0);       //Create a cell
	    cell1.setCellValue(new Date()); 
	    cell1.setCellStyle(style);
	    
	    
	    //to update the cell with current date in proper date and time format
		CellStyle style1=workbook.createCellStyle();
	    style1.setDataFormat(creationhelper.createDataFormat().getFormat("dd-mm-yyyy hh:mm:ss"));  //specify the data-time format
	    
	    
	    XSSFCell cell2 = sheet.createRow(2).createCell(0);       //Create a cell
	    cell2.setCellValue(new Date()); 
	    cell2.setCellStyle(style1);
		
		
		FileOutputStream fos  = new FileOutputStream ("./datafiles/dataformats.xlsx");  //Create a excel file
		workbook.write(fos);
		workbook.close();
		fos.close();
		System.out.println("DONE");

	}

}
