package project;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.Select;

public class T03EnterProductData {

	WebDriver driver;
	public T03EnterProductData(WebDriver driver) {
		this.driver=driver;
	}
	public void startdate(String date)
	{
		driver.findElement(By.id("startdate")).sendKeys(date);
	}
	public void insurancesum(Integer index)
	{
		new Select(driver.findElement(By.id("insurancesum"))).selectByIndex(index);
	}
	public void meritrating(Integer index)
	{
		new Select(driver.findElement(By.id("meritrating"))).selectByIndex(index);
	}
	public void damageinsurance(Integer index)
	{
		new Select(driver.findElement(By.id("damageinsurance"))).selectByIndex(index);
	}
	public void orignalProducts()
	{
		new Actions(driver).moveToElement(driver.findElement(By.id("LegalDefenseInsurance"))).click().perform();
	}
	public void courtesycar(String option)
	{
		new Select(driver.findElement(By.id("courtesycar"))).selectByVisibleText(option);
	}
	public void submitProductData() {
		driver.findElement(By.id("nextselectpriceoption")).click();
	}
}
