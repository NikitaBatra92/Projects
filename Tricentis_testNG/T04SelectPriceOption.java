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
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.interactions.Actions;

public class T04SelectPriceOption {

	WebDriver driver;
	String fpath = "D:\\Tricentis.xlsx";
	File file;
	FileInputStream fis;
	FileOutputStream fos;
	XSSFWorkbook wb;
	XSSFSheet sheet;
	XSSFRow row;
	XSSFCell cell;
	int rowIndex = 0;
	public T04SelectPriceOption(WebDriver driver) {
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
		
		
		if (testname == "automobile")
			rowIndex = 1;
		else if (testname == "truck")
			rowIndex = 6;
		else if (testname == "motorcycle")
			rowIndex = 11;
		else if (testname == "camper")
			rowIndex = 16;
		
		
		switch (type) {
		case "Silver":
			row = sheet.getRow(rowIndex);
			actPrice = driver.findElement(By.id("selectsilver_price")).getText();
			actClaim = driver.findElement(By.xpath("//table[@id='priceTable']/tbody/tr[2]/td[2]")).getText();
			actDis = driver.findElement(By.xpath("//table[@id='priceTable']/tbody/tr[3]/td[2]")).getText();
			actCover = driver.findElement(By.xpath("//table[@id='priceTable']/tbody/tr[4]/td[2]")).getText();
			if (expPrice.equals(actPrice))
				row.getCell(1).setCellValue("Price Matching Silver.");// System.out.println("Price Matching.");
			else
				row.getCell(1).setCellValue("Price Not Matching.");
			if (expClaim.equals(actClaim))
				row.getCell(2).setCellValue("Claim Matching.");
			else
				row.getCell(2).setCellValue("Claim not Matching.");
			if (expDis.equals(actDis))
				row.getCell(3).setCellValue("Discount Matching.");
			else
				row.getCell(3).setCellValue("Discount not Matching.");

			if (expCover.equals(actCover))
				row.getCell(4).setCellValue("Cover Matching.");
			else
				row.getCell(4).setCellValue("Cover not Matching.");
			new Actions(driver).moveToElement(driver.findElement(By.id("selectsilver"))).click().perform();
			driver.findElement(By.id("nextsendquote")).click();
			rowIndex++;
			break;

		case "Gold":			
			row = sheet.getRow(rowIndex);
			actPrice = driver.findElement(By.id("selectgold_price")).getText();
			actClaim = driver.findElement(By.xpath("//table[@id='priceTable']/tbody/tr[2]/td[3]")).getText();
			actDis = driver.findElement(By.xpath("//table[@id='priceTable']/tbody/tr[3]/td[3]")).getText();
			actCover = driver.findElement(By.xpath("//table[@id='priceTable']/tbody/tr[4]/td[3]")).getText();
			if (expPrice.equals(actPrice))
				row.getCell(1).setCellValue("Price Matching Gold.");// System.out.println("Price Matching.");
			else
				row.getCell(1).setCellValue("Price Not Matching.");
			if (expClaim.equals(actClaim))
				row.getCell(2).setCellValue("Claim Matching.");
			else
				row.getCell(2).setCellValue("Claim not Matching.");
			if (expDis.equals(actDis))
				row.getCell(3).setCellValue("Discount Matching.");
			else
				row.getCell(3).setCellValue("Discount not Matching.");

			if (expCover.equals(actCover))
				row.getCell(4).setCellValue("Cover Matching.");
			else
				row.getCell(4).setCellValue("Cover not Matching.");

			new Actions(driver).moveToElement(driver.findElement(By.id("selectgold"))).click().perform();
			driver.findElement(By.id("nextsendquote")).click();
			rowIndex++;
			break;
		case "Platinum":
			
			row = sheet.getRow(rowIndex);

			actPrice = driver.findElement(By.xpath("//table[@id='priceTable']/tbody/tr[1]/td[4]")).getText();
			actClaim = driver.findElement(By.xpath("//table[@id='priceTable']/tbody/tr[2]/td[4]")).getText();
			actDis = driver.findElement(By.xpath("//table[@id='priceTable']/tbody/tr[3]/td[4]")).getText();
			actCover = driver.findElement(By.xpath("//table[@id='priceTable']/tbody/tr[4]/td[4]")).getText();
			if (expPrice.equals(actPrice))
				row.getCell(1).setCellValue("Price Matching Platinum.");// System.out.println("Price Matching.");
			else
				row.getCell(1).setCellValue("Price Not Matching.");
			if (expClaim.equals(actClaim))
				row.getCell(2).setCellValue("Claim Matching.");
			else
				row.getCell(2).setCellValue("Claim not Matching.");
			if (expDis.equals(actDis))
				row.getCell(3).setCellValue("Discount Matching.");
			else
				row.getCell(3).setCellValue("Discount not Matching.");

			if (expCover.equals(actCover))
				row.getCell(4).setCellValue("Cover Matching.");
			else
				row.getCell(4).setCellValue("Cover not Matching.");

			new Actions(driver).moveToElement(driver.findElement(By.id("selectplatinum"))).click().perform();
			driver.findElement(By.id("nextsendquote")).click();
			rowIndex++;
			break;
			
		case "Ultimate":
		
			row = sheet.getRow(rowIndex);
			actPrice = driver.findElement(By.id("selectultimate_price")).getText();
			actClaim = driver.findElement(By.xpath("//table[@id='priceTable']/tbody/tr[2]/td[5]")).getText();
			actDis = driver.findElement(By.xpath("//table[@id='priceTable']/tbody/tr[3]/td[5]")).getText();
			actCover = driver.findElement(By.xpath("//table[@id='priceTable']/tbody/tr[4]/td[5]")).getText();
			if (expPrice.equals(actPrice))
				row.getCell(1).setCellValue("Price Matching Ultimate.");// System.out.println("Price Matching.");
			else
				row.getCell(1).setCellValue("Price Not Matching.");
			if (expClaim.equals(actClaim))
				row.getCell(2).setCellValue("Claim Matching.");
			else
				row.getCell(2).setCellValue("Claim not Matching.");
			if (expDis.equals(actDis))
				row.getCell(3).setCellValue("Discount Matching.");
			else
				row.getCell(3).setCellValue("Discount not Matching.");

			if (expCover.equals(actCover))
				row.getCell(4).setCellValue("Cover Matching.");
			else
				row.getCell(4).setCellValue("Cover not Matching.");

			new Actions(driver).moveToElement(driver.findElement(By.id("selectultimate"))).click().perform();
			driver.findElement(By.id("nextsendquote")).click();
			rowIndex++;
			break;
		}
		wb.write(fos);
		wb.close();
		fis.close();

	}

}
