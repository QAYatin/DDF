package testCase;

import java.io.File;
import java.util.Properties;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import utility.GeneralUtils;

public class Tc_001_HRM_Login 
{
	WebDriver driver;
	XSSFWorkbook workbook;
	XSSFSheet sheet;
	XSSFCell cell;
	
	 @BeforeTest
		public void start() throws Exception
		{
		 	GeneralUtils util = new GeneralUtils();
		 	Properties po = util.readPropertiesFile("C:\\Users\\Yattu\\eclipse-workspace\\Nivi_Workspace\\DDF\\TestData\\HRM.properties");
		 	System.setProperty("webdriver.chrome.driver","C:\\Users\\Yattu\\OneDrive\\Documents\\Jar files\\chrome\\chromedriver.exe");
			driver = new ChromeDriver();
			driver.get(po.getProperty("orangeHRM"));
			driver.manage().window().maximize();
		}
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
				workbook.close();
			}
	 
	 
	 @Test
	  public void logic() throws Exception
	  {
		 GeneralUtils util = new GeneralUtils();
		 Properties po = util.readPropertiesFile("C:\\Users\\Yattu\\eclipse-workspace\\Nivi_Workspace\\DDF\\TestData\\HRM.properties");
		 setReadExcelFile("C:\\Users\\Yattu\\eclipse-workspace\\Nivi_Workspace\\DDF\\TestData\\OrangeHRMLogin.xlsx", "HRMLogin");
		 
		 for(int i=1; i<=getRowSize(); i++)
		 {
			 driver.findElement(By.xpath(po.getProperty("username"))).clear();
			 driver.findElement(By.xpath(po.getProperty("username"))).sendKeys(getCellData(i,2));
			 Thread.sleep(1000);
			 driver.findElement(By.xpath(po.getProperty("password"))).clear();
			 driver.findElement(By.xpath(po.getProperty("password"))).sendKeys(getCellData(i,3));
			 Thread.sleep(1000);
			 driver.findElement(By.xpath(po.getProperty("loginButton"))).click();
			 Thread.sleep(1000);
			 
			 int locSize = driver.findElements(By.xpath(po.getProperty("loggedInUsername"))).size();
			 
			 if(locSize>0)
			 {
				 driver.findElement(By.xpath(po.getProperty("loggedInUsername"))).click();
				 Thread.sleep(1000);
				 driver.findElement(By.xpath(po.getProperty("logout"))).click();
				 System.out.println("Login Passed");
			 }
			 else
			 {
				 System.out.println("Login Failed");
			 }
		 }
		 closeWorkbook();
	  }
	 
	 
	 
	 @AfterTest
		public void endTest() throws Exception
		{
			Thread.sleep(5000);
			driver.close();
		}
}
