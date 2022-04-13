package datadriven;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.ITestResult;

public class XLUtility extends LoginTest {
  static FileInputStream fi;
  
  static FileOutputStream fo;
  
  static XSSFWorkbook workbook;
  
  static XSSFSheet sheet;
  
  static XSSFRow row;
  
  static XSSFCell cell;
  
  static CellStyle style;
  
  static String path;
  
  XLUtility(String path) {
    XLUtility.path = path;
  }
  
  public int getRowCount(String sheetName) throws IOException {
    fi = new FileInputStream(path);
    workbook = new XSSFWorkbook(fi);
    sheet = workbook.getSheet(sheetName);
    int rowcount = sheet.getLastRowNum();
    workbook.close();
    fi.close();
    return rowcount;
  }
  
  public int getCellCount(String sheetName, int rownum) throws IOException {
    fi = new FileInputStream(path);
    workbook = new XSSFWorkbook(fi);
    sheet = workbook.getSheet(sheetName);
    row = sheet.getRow(rownum);
    int cellcount = row.getLastCellNum();
    workbook.close();
    fi.close();
    return cellcount;
  }
  
  public String getCellData(String sheetName, int rownum, int colnum) throws IOException {
    String data;
    fi = new FileInputStream(path);
    workbook = new XSSFWorkbook(fi);
    sheet = workbook.getSheet(sheetName);
    row = sheet.getRow(rownum);
    cell = row.getCell(colnum);
    DataFormatter formatter = new DataFormatter();
    try {
      data = formatter.formatCellValue((Cell)cell);
    } catch (Exception e) {
      data = "";
    } 
    workbook.close();
    fi.close();
    return data;
  }
  
  public String setCellData(String sheetName, int rownum, int colnum, String data) throws IOException {
    File xlfile = new File(path);
    if (!xlfile.exists()) {
      workbook = new XSSFWorkbook();
      fo = new FileOutputStream(path);
      workbook.write(fo);
    } 
    fi = new FileInputStream(path);
    workbook = new XSSFWorkbook(fi);
    if (workbook.getSheetIndex(sheetName) == -1)
      workbook.createSheet(sheetName); 
    sheet = workbook.getSheet(sheetName);
    if (sheet.getRow(rownum) == null)
      sheet.createRow(rownum); 
    row = sheet.getRow(rownum);
    cell = row.createCell(colnum);
    cell.setCellValue(data);
    fo = new FileOutputStream(path);
    workbook.write(fo);
    workbook.close();
    fi.close();
    fo.close();
    return data;
  }
  
  public static void fillGreenColor(String sheetName, int rownum, int colnum) throws IOException {
    fi = new FileInputStream(path);
    workbook = new XSSFWorkbook(fi);
    sheet = workbook.getSheet(sheetName);
    row = sheet.getRow(rownum);
    cell = row.getCell(colnum);
    style = (CellStyle)workbook.createCellStyle();
    style.setFillForegroundColor(IndexedColors.GREEN.getIndex());
    style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    cell.setCellStyle(style);
    workbook.write(fo);
    workbook.close();
    fi.close();
    fo.close();
  }
  
  public static void fillRedColor(String sheetName, int rownum, int colnum) throws IOException {
    fi = new FileInputStream(path);
    workbook = new XSSFWorkbook(fi);
    sheet = workbook.getSheet(sheetName);
    row = sheet.getRow(rownum);
    cell = row.getCell(colnum);
    style = (CellStyle)workbook.createCellStyle();
    style.setFillForegroundColor(IndexedColors.RED.getIndex());
    style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    cell.setCellStyle(style);
    workbook.write(fo);
    workbook.close();
    fi.close();
    fo.close();
  }
  
  public void updateTestResult(String excellocation, String sheetName, String testCaseName, String testStatus) throws IOException {
    try {
      FileInputStream file = new FileInputStream(new File(excellocation));
      XSSFWorkbook workbook = new XSSFWorkbook(file);
      XSSFSheet sheet = workbook.getSheet(sheetName);
      int totalRow = sheet.getLastRowNum() + 1;
      for (int i = 1; i < totalRow; i++) {
        XSSFRow r = sheet.getRow(i);
        String ce = r.getCell(1).getStringCellValue();
        if (ce.contains(testCaseName)) {
          r.createCell(2).setCellValue(testStatus);
          file.close();
          System.out.println("resule updated");
          FileOutputStream outFile = new FileOutputStream(new File(excellocation));
          workbook.write(outFile);
          outFile.close();
          break;
        } 
      } 
    } catch (Exception e) {
      e.printStackTrace();
    } 
  }
  
  public void updateResult(String path, String sheetName, ITestResult result) throws IOException {
    FileInputStream file = new FileInputStream(new File(path));
    XSSFWorkbook workbook = new XSSFWorkbook(file);
    XSSFSheet sheet = workbook.getSheet(sheetName);
    XSSFRow r = sheet.getRow(count);
    if (result.getStatus() == 1) {
      r.createCell(3).setCellValue("PASS");
      cell = r.getCell(3);
      style = (CellStyle)workbook.createCellStyle();
      style.setFillForegroundColor(IndexedColors.GREEN.getIndex());
      style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
      cell.setCellStyle(style);
    } else if (result.getStatus() == 2) {
      r.createCell(3).setCellValue("FAIL");
      cell = r.getCell(3);
      style = (CellStyle)workbook.createCellStyle();
      style.setFillForegroundColor(IndexedColors.RED.getIndex());
      style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
      cell.setCellStyle(style);
    } else if (result.getStatus() == 3) {
      r.createCell(3).setCellValue("SKIPPED");
      cell = r.getCell(3);
      style = (CellStyle)workbook.createCellStyle();
      style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
      style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
      cell.setCellStyle(style);
    } 
    file.close();
    fo = new FileOutputStream(path);
    workbook.write(fo);
    workbook.close();
  }
}
