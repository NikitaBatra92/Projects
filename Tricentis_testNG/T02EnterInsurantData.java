package project;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.Select;

public class T02EnterInsurantData {

	WebDriver driver;

	public T02EnterInsurantData(WebDriver driver) {
		this.driver = driver;
	}

	public void firstname(String fname) {
		driver.findElement(By.id("firstname")).sendKeys(fname);
	}

	public void lastname(String lname) {
		driver.findElement(By.id("lastname")).sendKeys(lname);
	}

	public void birthdate(String dob) {
		driver.findElement(By.id("birthdate")).sendKeys("04/20/1992");
	}

	public void gender() {
		Actions act = new Actions(driver);
		act.moveToElement(driver.findElement(By.id("genderfemale"))).click().perform();
	}

	public void streetaddress(String address) {
		driver.findElement(By.id("streetaddress")).sendKeys(address);
	}

	public void country(String country) {
		new Select(driver.findElement(By.id("country"))).selectByVisibleText(country);
	}

	public void zipcode(String code) {
		driver.findElement(By.id("zipcode")).sendKeys(code);
	}

	public void city(String city) {
		driver.findElement(By.id("city")).sendKeys(city);
	}

	public void occupation(String occup) {
		new Select(driver.findElement(By.id("occupation"))).selectByVisibleText(occup);
	}

	public void hobbies(boolean speeding, boolean bungee, boolean diving, boolean skydiving, boolean other) {
		if (speeding)
			new Actions(driver).moveToElement(driver.findElement(By.xpath(
					"/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[1]/form[1]/div[1]/section[2]/div[10]/p[1]/label[1]/span[1]")))
					.click().perform();
		if (bungee)
			new Actions(driver).moveToElement(driver.findElement(By.xpath(
					"/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[1]/form[1]/div[1]/section[2]/div[10]/p[1]/label[2]/span[1]")))
					.click().perform();
		if (diving)
			new Actions(driver).moveToElement(driver.findElement(By.xpath(
					"/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[1]/form[1]/div[1]/section[2]/div[10]/p[1]/label[3]/span[1]")))
					.click().perform();
		if (skydiving)
			new Actions(driver).moveToElement(driver.findElement(By.xpath(
					"/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[1]/form[1]/div[1]/section[2]/div[10]/p[1]/label[4]/span[1]")))
					.click().perform();
		if (other)
			new Actions(driver).moveToElement(driver.findElement(By.xpath(
					"/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[1]/form[1]/div[1]/section[2]/div[10]/p[1]/label[5]/span[1]")))
					.click().perform();
		else
			new Actions(driver).moveToElement(driver.findElement(By.xpath(
					"/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[1]/form[1]/div[1]/section[2]/div[10]/p[1]/label[5]/span[1]")))
					.click().perform();
	}

	public void submitInsurantData() {
		driver.findElement(By.id("nextenterproductdata")).click();
	}
}
