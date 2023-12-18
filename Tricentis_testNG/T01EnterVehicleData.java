package project;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.support.ui.Select;

public class T01EnterVehicleData {

	WebDriver driver;

	public T01EnterVehicleData(WebDriver driver) {
		this.driver = driver;
	}

	public void selectMake(String make) {
		new Select(driver.findElement(By.id("make"))).selectByVisibleText(make);
	}

	public void setEnginePerformance(String ep) {
		driver.findElement(By.id("engineperformance")).sendKeys(ep);
	}

	public void setManDate(String date) {
		driver.findElement(By.id("dateofmanufacture")).sendKeys(date);
	}

	public void seats(String no) {
		new Select(driver.findElement(By.id("numberofseats"))).selectByVisibleText(no);
	}
public void seatsmotorcycle(String no) {
	new Select(driver.findElement(By.id("numberofseatsmotorcycle"))).selectByVisibleText(no);
}
	public void fuel(String fueltype) {
		new Select(driver.findElement(By.id("fuel"))).selectByVisibleText(fueltype);
	}

	public void listprice(String price) {
		driver.findElement(By.id("listprice")).sendKeys(price);
	}

	public void licensenumber(String no) {
		driver.findElement(By.id("licenseplatenumber")).sendKeys(no);
	}

	public void annualmileage(String mileage) {
		driver.findElement(By.id("annualmileage")).sendKeys(mileage);
	}

	public void submitVehicleData() {
		driver.findElement(By.id("nextenterinsurantdata")).click();
	}
	public void payload(String load) {
	driver.findElement(By.id("payload")).sendKeys(load);}
	public void totalweight(String weight) {
	driver.findElement(By.id("totalweight")).sendKeys(weight);}
	public void modelMotorcycle(Integer model) {
	new Select(driver.findElement(By.id("model"))).selectByIndex(model);}
	public void cylindercapacity(String capacity) {
	driver.findElement(By.id("cylindercapacity")).sendKeys(capacity);}
}










