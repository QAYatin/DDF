package utility;

import java.io.File;
import java.io.FileOutputStream;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtils 
{
	XSSFWorkbook workbook;
	XSSFSheet sheet;
	XSSFCell cell;
	
	XSSFWorkbook writeWorkbook;
	XSSFSheet writeSheet;
	XSSFCell writeCell;
	
	
	public void setReadExcelFile(String filePath,String sheetName) throws Exception
	{
		File fi = new File(filePath);
		workbook = new XSSFWorkbook(fi);
		sheet = workbook.getSheet(sheetName);
	}
	
	public String getCellData(int rowNum,int colNum)
	{
		cell = sheet.getRow(rowNum).getCell(colNum);
		return cell.getStringCellValue();
	}
	
	public int getRowSize()
	{
		int rowSize = sheet.getLastRowNum();
		return rowSize;
	}
	
	public void closeWorkbook() throws Exception
	{
		
	}
	
	public void setWriteExcelFile(String filePath,String sheetName) throws Exception
	{
		File fi = new File(filePath);
		writeWorkbook = new XSSFWorkbook(fi);
		writeSheet = writeWorkbook.getSheet(sheetName);
	}
	
	public void setCellData(int rowNum,int colNum, String inputText)
	{
		writeSheet.getRow(rowNum).getCell(colNum).setCellValue(inputText);
	}
	
	public void closeWriteWorkbook(String path) throws Exception
	{
		FileOutputStream fo = new FileOutputStream(path);
		writeWorkbook.write(fo);
		writeWorkbook.close();
	}

}
