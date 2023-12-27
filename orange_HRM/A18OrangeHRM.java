package com.Assignments;

import org.testng.annotations.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.DataProvider;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.AfterTest;

public class A18OrangeHRM {
	WebDriver driver;
	String fpath = "D:\\OHRMAssignment_Data.xlsx";
	File file;
	FileInputStream fis;
	FileOutputStream fos;
	XSSFWorkbook wb;
	XSSFSheet sheet;
	XSSFRow row;
	XSSFCell cell;
	A18OrangeHRMlogin A1;
	int count=1;
	int rows;  	
	@Test
	public void test1login() throws InterruptedException {	
		//rows=sheet.getPhysicalNumberOfRows();
		for(int i=1;i<3;i++)
		  {
			  row= sheet.getRow(i);
		A1.setUserName("Admin");
		A1.setPassword("admin123");
		A1.login();
		// 1. Login with Admin
		driver.findElement(By.linkText("PIM")).click();
		// 2. Click on PIM
		driver.findElement(By.xpath("/html/body/div/div[1]/div[2]/div[2]/div/div[2]/div[1]/button")).click();
		// 3. Click on + Add button

		//row= sheet.getRow(count);
		//row = sheet.getRow(1);
		cell = row.getCell(1);
		// 4. Enter First Name
		driver.findElement(By.name("firstName")).sendKeys(cell.getStringCellValue());
		cell = row.getCell(2);
		// 5. Enter Middle Name
		driver.findElement(By.name("middleName")).sendKeys(cell.getStringCellValue());
		cell = row.getCell(3);
		// 6. Enter Last Name
		driver.findElement(By.name("lastName")).sendKeys(cell.getStringCellValue());
		// 7. Click on Create Login Details
		driver.findElement(
				By.xpath("/html/body/div/div[1]/div[2]/div[2]/div/div/form/div[1]/div[2]/div[2]/div/label/span"))
				.click();
		// 8. Enter User Name
		cell = row.getCell(4);
		driver.findElement(By.xpath("/html[1]/body[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/form[1]/div[1]"
				+ "/div[2]/div[3]/div[1]/div[1]/div[1]/div[2]/input[1]")).sendKeys(cell.getStringCellValue());
		// 9. Enter Password
		// 10. Enter Confirm Password
		cell = row.getCell(5);

		driver.findElement(By.xpath("/html[1]/body[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/form[1]/div[1]"
				+ "/div[2]/div[4]/div[1]/div[1]/div[1]/div[2]/input[1]")).sendKeys(cell.getStringCellValue());
		driver.findElement(By.xpath("/html[1]/body[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/form[1]/div[1]"
				+ "/div[2]/div[4]/div[1]/div[2]/div[1]/div[2]/input[1]")).sendKeys(cell.getStringCellValue());
		// 11. Click on Save Button
		driver.findElement(By.xpath("/html/body/div/div[1]/div[2]/div[2]/div/div/form/div[2]/button[2]")).click();
		Thread.sleep(5000);
		// 12. Logout
		A1.logout();

		// 13. Login using data provided in step no 8 & 9
		A1.setUserName(row.getCell(4).getStringCellValue());
		A1.setPassword(row.getCell(5).getStringCellValue());
		A1.login();
		driver.findElement(By.xpath("/html/body/div/div[1]/div[1]/aside/nav/div[2]/ul/li[3]/a")).click();
		// 14. A. Read the employee ID & store in the Excel file under Emp ID Column.
				WebElement empidbox = driver
						.findElement(By.xpath("/html/body/div/div[1]/div[2]/div[2]/div/div/div/div[2]/div[1]/form"
								+ "/div[2]/div[1]/div[1]/div/div[2]/input"));
				String empid = empidbox.getAttribute("value");
				cell = row.getCell(6);
				cell.setCellValue(empid);

				// 14. B. Read the name displayed at Right Top corner of page and store the same
				// in Excel sheet under Expected Message.
				WebElement result = driver
						.findElement(By.xpath("/html/body/div/div[1]/div[1]/header/div[1]/div[2]/ul/li/span/p"));
				cell = row.getCell(8);//saved in actual message cell.
				cell.setCellValue(result.getText());
				// 15. Compare Expected Message and Actual Message from Excel sheet
				if (row.getCell(7).getStringCellValue().equals(row.getCell(8).getStringCellValue())) 
				{
					// 16. Mark Result column as Pass or Fail accordingly
					cell = row.getCell(9);
					cell.setCellValue("Pass");
				} else {
					cell = row.getCell(9);
					cell.setCellValue("Fail");
				}
				A1.logout();
				Thread.sleep(2000);
				// 18. Login with Admin
				A1.setUserName("Admin");
				A1.setPassword("admin123");
				A1.login();
				// 19. Click on PIM
				driver.findElement(By.linkText("PIM")).click();
				// 20. Enter First Name(Step No 4) in Employee Name textbox
				driver.findElement(By.xpath("/html/body/div/div[1]/div[2]/div[2]/div/div[1]/div[2]/form"
						+ "/div[1]/div/div[1]/div/div[2]/div/div/input")).sendKeys(row.getCell(1).getStringCellValue());
				// 21. Click on Search
				driver.findElement(By.xpath("//button[@type='submit']")).click();
				Thread.sleep(3000);
				JavascriptExecutor js = (JavascriptExecutor)driver;
				// 22. Respective user information will be displayed. Delete this data.
				js.executeScript("arguments[0].click()", driver.findElement(By.xpath("/html/body/div/div[1]/div[2]/div[2]/div/div[2]/div[3]/div/div[2]/div/div/div[9]/div/button[1]/i")));
				//driver.findElement(By.xpath("/html/body/div/div[1]/div[2]/div[2]/div/div[2]/div[3]/div/div[2]/div/div/div[9]/div/button[1]/i")).click();
               // Actions act= new Actions(driver); 
                //act.moveToElement(driver.findElement(By.xpath("/html/body/div/div[3]/div/div/div/div[3]/button[2]"))).click().perform();
				//driver.findElement(By.xpath("/html/body/div/div[3]/div/div/div/div[3]/button[2]")).click();
				js.executeScript("arguments[0].click()", driver.findElement(By.xpath("/html/body/div/div[3]/div/div/div/div[3]/button[2]")));
				Thread.sleep(5000);
				// 23. Logout
				A1.logout();
		  }
	}
	
	@BeforeTest
	public void beforeTest() throws IOException {
		file = new File(fpath);
		fis = new FileInputStream(file);
		wb = new XSSFWorkbook(fis);
		sheet = wb.getSheetAt(0);
		fos = new FileOutputStream(file);
		driver = new ChromeDriver();
		driver.get("https://opensource-demo.orangehrmlive.com/web/index.php/auth/login");
		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(20));
		A1 = new A18OrangeHRMlogin(driver);
	}

	@AfterTest
	public void afterTest() throws IOException {
		wb.write(fos);
		wb.close();
		fis.close();
		driver.close();
	}

}
