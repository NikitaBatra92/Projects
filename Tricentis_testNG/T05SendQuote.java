package project;

import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;

public class T05SendQuote {
	WebDriver driver;

	public T05SendQuote(WebDriver driver) {
		this.driver = driver;
	}

	public void email(String emailid) {
		driver.findElement(By.id("email")).sendKeys(emailid);
	}

	public void username(String name) {
		driver.findElement(By.id("username")).sendKeys(name);
	}

	public void password(String pass) {
		driver.findElement(By.id("password")).sendKeys(pass);
	}

	public void confirmpassword(String pass) {
		driver.findElement(By.id("confirmpassword")).sendKeys(pass);
	}

	public void comment() {
		driver.findElement(By.id("Comments")).sendKeys("Comments");
	}

	public void sendemail() throws InterruptedException {
		JavascriptExecutor js = (JavascriptExecutor)driver;
		js.executeScript("arguments[0].click()", driver.findElement(By.xpath("//button[@id='sendemail']")));
		
		//driver.findElement(By.id("sendemail")).click();
		Thread.sleep(9000);
	}

	public void message() {
		//System.out.println("Message : " + driver.findElement(By.xpath("/html[1]/body[1]/div[4]/h2[1]")).getText());
		driver.findElement(By.className("confirm")).click();
	}
}
