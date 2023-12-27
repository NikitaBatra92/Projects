package com.Assignments;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;

public class A18OrangeHRMlogin {
	WebDriver driver;
	public A18OrangeHRMlogin(WebDriver d)
	{
		driver = d;
	}
	public void setUserName(String name)
	{
		driver.findElement(By.name("username")).sendKeys(name);
	}
	public void setPassword(String pass)
	{
		driver.findElement(By.name("password")).sendKeys(pass);
	}
	public void login() {
		
		driver.findElement(By.xpath("//button[@type='submit']")).click();
	}
	public void logout()
	{
		String actUrl=driver.getCurrentUrl();
		String expUrl="https://opensource-demo.orangehrmlive.com/web/index.php/auth/login";
		if(actUrl!=(expUrl))
		  {  
			driver.findElement(By.xpath("/html[1]/body[1]/div[1]/div[1]/div[1]/header[1]/div[1]/div[2]/ul[1]/li[1]/span[1]/i[1]")).click();
		    driver.findElement(By.linkText("Logout")).click(); 
		  } 
	}
}
