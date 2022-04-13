package exceloperations;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

// Workbook-->Sheet-->Rows->Cells

public class WrtingExcelDemo1 {

	public static void main(String[] args) throws IOException {
	
		XSSFWorkbook workbook=new XSSFWorkbook();           //create an empty workbook
		XSSFSheet sheet=workbook.createSheet("Emp Info");   //create a new sheet inside the workbook
		
		Object empdata[][]= {	{"EmpID","Name","Job"},     //store the data values in a heterogenous 2D array of type 'Object' to store any kind of data
								{101,"David","Enginner"},
								{102,"Smith","Manager"},
								{103,"Scott","Analyst"}
							};
		
		//Using for loop
		
   /*   
        int rows=empdata.length;                      //will return the no. of rows in 2D array
		int cols=empdata[0].length;                   //will return the no. of columns in 2D array
		
		System.out.println(rows);                     //4
		System.out.println(cols);                     //3
		
		for(int r=0;r<rows;r++)                       //outer for loop for rows
		{
			XSSFRow row=sheet.createRow(r);           //create a row in the sheet
			
			for(int c=0;c<cols;c++)                   //inner for loop for columns/cells
			{
				XSSFCell cell=row.createCell(c);      //create cells in the rows
				Object value=empdata[r][c];           //get all the data from 2D array according to rows & columns
				
				if(value instanceof String)         
					cell.setCellValue((String)value);  //convert data from object to string and then setting the string data value in the cells
				if(value instanceof Integer)
					cell.setCellValue((Integer)value); //convert data from object to integer and then setting the integer data value in the cells
				if(value instanceof Boolean)
					cell.setCellValue((Boolean)value);
				
			}
		}
	*/
		
		/// using for...each loop
		
		int rowCount=0;
		
		for(Object emp[]:empdata)                                // create a array variable i.e emp[] of type Object and store the 2D array data of one row 
		{
			XSSFRow row=sheet.createRow(rowCount++);             //create a row and then increment the rowcount value by 1
			
			int columnCount=0;		
				for(Object value:emp)                             //getting every value from 'emp' variable and storing it in 'value' variable for updating/writing in the cells
				{
					XSSFCell cell=row.createCell(columnCount++);  //create a cell and then increment the columncount value by 1
					
					if(value instanceof String)                    // to check whether the cell type is string , numeric or boolean
							cell.setCellValue((String)value);
					if(value instanceof Integer)
							cell.setCellValue((Integer)value);
					if(value instanceof Boolean)
						cell.setCellValue((Boolean)value);	
							
				}
		}
		
		
		
		
		String filePath=".\\datafiles\\employee.xlsx";                //path where the above code will create the excel file
		FileOutputStream outstream=new FileOutputStream(filePath);    //create and open the file in the fileoutputstream mode so that we could write data in file
		workbook.write(outstream);                                    //write or attach the workbook into the excel file 
		
		outstream.close();                                            //close outstream mode
		
		System.out.println("Employee.xlsx file written successfully...");
	}

}
