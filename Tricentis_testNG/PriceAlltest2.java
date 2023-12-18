
package project;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.interactions.Actions;

public class PriceAlltest2 {

	WebDriver driver;
	String fpath = "D:\\tricentisnew.xlsx";
	File file;
	FileInputStream fis;
	FileOutputStream fos;
	XSSFWorkbook wb;
	XSSFSheet sheet;
	XSSFRow row;
	XSSFCell cell;
int rowIndex=1;
	public PriceAlltest2(WebDriver driver) {
		this.driver = driver;
	}

	public void checkPrice(String expPrice, String expClaim, String expDis, String expCover, String type,
			String testname) throws IOException {
		String actPrice, actClaim, actDis, actCover;
		file = new File(fpath);
		fis = new FileInputStream(file);
		wb = new XSSFWorkbook(fis);
		sheet = wb.getSheetAt(0);
		fos = new FileOutputStream(file);

		

		switch (type) {
		case "Silver":
			if (testname == "Automobile")
				sheet=wb.getSheetAt(0);
			else if (testname == "Truck")
				sheet=wb.getSheetAt(1);
			else if (testname == "Motorcycle")
				sheet=wb.getSheetAt(2);
			else if (testname == "Camper")
				sheet=wb.getSheetAt(3);
			row = sheet.getRow(1);
			actPrice = driver.findElement(By.id("selectsilver_price")).getText();
			actClaim = driver.findElement(By.xpath("//table[@id='priceTable']/tbody/tr[2]/td[2]")).getText();
			actDis = driver.findElement(By.xpath("//table[@id='priceTable']/tbody/tr[3]/td[2]")).getText();
			actCover = driver.findElement(By.xpath("//table[@id='priceTable']/tbody/tr[4]/td[2]")).getText();
			if (expPrice.equals(actPrice) && expClaim.equals(actClaim) && expDis.equals(actDis)
					&& expCover.equals(actCover))
			sheet.getRow(1).getCell(5).setCellValue("Pass");
			else
				sheet.getRow(1).getCell(5).setCellValue("Fail");
System.out.println(testname+":"+type+":"+actPrice+","+actClaim+","+actDis+","+actCover);
			new Actions(driver).moveToElement(driver.findElement(By.id("selectsilver"))).click().perform();
			driver.findElement(By.id("nextsendquote")).click();
		
			break;

		case "Gold":
			if (testname == "Automobile")
				sheet=wb.getSheetAt(0);
			else if (testname == "Truck")
				sheet=wb.getSheetAt(1);
			else if (testname == "Motorcycle")
				sheet=wb.getSheetAt(2);
			else if (testname == "Camper")
				sheet=wb.getSheetAt(3);
			row = sheet.getRow(2);

			actPrice = driver.findElement(By.id("selectgold_price")).getText();
			actClaim = driver.findElement(By.xpath("//table[@id='priceTable']/tbody/tr[2]/td[3]")).getText();
			actDis = driver.findElement(By.xpath("//table[@id='priceTable']/tbody/tr[3]/td[3]")).getText();
			actCover = driver.findElement(By.xpath("//table[@id='priceTable']/tbody/tr[4]/td[3]")).getText();
			if (expPrice.equals(actPrice) && expClaim.equals(actClaim) && expDis.equals(actDis)
					&& expCover.equals(actCover))
				sheet.getRow(2).getCell(5).setCellValue("Pass");
			else
				sheet.getRow(2).getCell(5).setCellValue("Fail");
			System.out.println(testname+":"+type+":"+actPrice+","+actClaim+","+actDis+","+actCover);
			new Actions(driver).moveToElement(driver.findElement(By.id("selectgold"))).click().perform();
			driver.findElement(By.id("nextsendquote")).click();
		
			break;
		case "Platinum":
			if (testname == "Automobile")
				sheet=wb.getSheetAt(0);
			else if (testname == "Truck")
				sheet=wb.getSheetAt(1);
			else if (testname == "Motorcycle")
				sheet=wb.getSheetAt(2);
			else if (testname == "Camper")
				sheet=wb.getSheetAt(3);
			row = sheet.getRow(3);

			actPrice = driver.findElement(By.xpath("//table[@id='priceTable']/tbody/tr[1]/td[4]")).getText();
			actClaim = driver.findElement(By.xpath("//table[@id='priceTable']/tbody/tr[2]/td[4]")).getText();
			actDis = driver.findElement(By.xpath("//table[@id='priceTable']/tbody/tr[3]/td[4]")).getText();
			actCover = driver.findElement(By.xpath("//table[@id='priceTable']/tbody/tr[4]/td[4]")).getText();
			if (expPrice.equals(actPrice) && expClaim.equals(actClaim) && expDis.equals(actDis)
					&& expCover.equals(actCover))
				sheet.getRow(3).getCell(5).setCellValue("Pass");
			else
				sheet.getRow(3).getCell(5).setCellValue("Fail");
			System.out.println(testname+":"+type+":"+actPrice+","+actClaim+","+actDis+","+actCover);
			new Actions(driver).moveToElement(driver.findElement(By.id("selectplatinum"))).click().perform();
			driver.findElement(By.id("nextsendquote")).click();
		
			break;

		case "Ultimate":
			if (testname == "Automobile")
				sheet=wb.getSheetAt(0);
			else if (testname == "Truck")
				sheet=wb.getSheetAt(1);
			else if (testname == "Motorcycle")
				sheet=wb.getSheetAt(2);
			else if (testname == "Camper")
				sheet=wb.getSheetAt(3);
			row = sheet.getRow(4);
			actPrice = driver.findElement(By.id("selectultimate_price")).getText();
			actClaim = driver.findElement(By.xpath("//table[@id='priceTable']/tbody/tr[2]/td[5]")).getText();
			actDis = driver.findElement(By.xpath("//table[@id='priceTable']/tbody/tr[3]/td[5]")).getText();
			actCover = driver.findElement(By.xpath("//table[@id='priceTable']/tbody/tr[4]/td[5]")).getText();
			if (expPrice.equals(actPrice) && expClaim.equals(actClaim) && expDis.equals(actDis)
					&& expCover.equals(actCover))
				sheet.getRow(4).getCell(5).setCellValue("Pass");
			else
				sheet.getRow(4).getCell(5).setCellValue("Fail");
			System.out.println(testname+":"+type+":"+actPrice+","+actClaim+","+actDis+","+actCover);
			new Actions(driver).moveToElement(driver.findElement(By.id("selectultimate"))).click().perform();
			driver.findElement(By.id("nextsendquote")).click();

			break;
		}
		wb.write(fos);
		wb.close();
		fis.close();

	}

}
