package project;
import org.testng.annotations.Test;


import java.io.IOException;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.Calendar;
import java.util.Date;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.DataProvider;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.AfterTest;

public class MainAllTests {
	WebDriver driver;
	T01EnterVehicleData vd;
	T02EnterInsurantData id;
	T03EnterProductData pd;
	T04SelectPriceOption cp;
	T05SendQuote sq;
	
	@Test(priority=1,dataProvider = "automobileData")
	public void automobileTest(String price, String claim, String dis, String cover, String type,String testname) throws InterruptedException, IOException {
		driver.findElement(By.partialLinkText("Auto")).click();
		
		//Enter Vehicle Data(Automobile)
		vd.selectMake("Volkswagen");
		vd.setEnginePerformance("1900");
		Calendar cal = Calendar.getInstance();
		cal.add(Calendar.DATE, -1);
		SimpleDateFormat s = new SimpleDateFormat("MM/dd/yyyy");
		String manDate = s.format(new Date(cal.getTimeInMillis()));
		vd.setManDate(manDate);
		vd.seats("4");
		vd.fuel("Petrol");
		vd.listprice("7000");
		vd.licensenumber("MH27AF5194");
		vd.annualmileage("5000");
		vd.submitVehicleData();
		
		//Enter Insurant data
		id.firstname("Nikita");
		id.lastname("Batra");
		id.birthdate("04/20/1992");
		id.gender();
		id.streetaddress("Bavdhan");
		id.country("India");
		id.zipcode("411008");
		id.city("Pune");
		id.occupation("Employee");
		id.hobbies(true,true,false,false,false);
		id.submitInsurantData();
		
		//Enter Product Data
		Calendar.getInstance();
		cal.add(Calendar.DATE, 33);
		SimpleDateFormat s1 = new SimpleDateFormat("MM/dd/yyyy");
		String futureDate = s1.format(new Date(cal.getTimeInMillis()));
		pd.startdate(futureDate);
		pd.insurancesum(1);
		pd.meritrating(2);
		pd.damageinsurance(1);
		pd.orignalProducts();
		pd.courtesycar("Yes");
		pd.submitProductData();
		cp.checkPrice(price, claim, dis, cover, type,testname);		
		sq.email("nikitaokeshwani@gmail.com");
		sq.username("Nikita");
		sq.password("Batra@123");
		sq.confirmpassword("Batra@123");
		sq.comment();
		sq.sendemail();	
		sq.message();
	}
	@Test(priority=2,dataProvider = "truckData")
	public void truckTest(String price, String claim, String dis, String cover, String type,String testname) throws InterruptedException, IOException {
		driver.findElement(By.partialLinkText("Truck")).click();
		//Enter Vehicle Data(Truck)
				vd.selectMake("Volvo");
				vd.setEnginePerformance("1900");
				Calendar cal = Calendar.getInstance();
				cal.add(Calendar.DATE, -1);
				SimpleDateFormat s = new SimpleDateFormat("MM/dd/yyyy");
				String manDate = s.format(new Date(cal.getTimeInMillis()));
				vd.setManDate(manDate);
				vd.seats("4");
				vd.fuel("Petrol");
				vd.payload("1000");
				vd.totalweight("25000");
				vd.listprice("7000");
				vd.licensenumber("MH27AF5194");
				vd.annualmileage("5000");
				vd.submitVehicleData();
				//Enter Insurant data
				id.firstname("Nikita");
				id.lastname("Batra");
				id.birthdate("04/20/1992");
				id.gender();
				id.streetaddress("Bavdhan");
				id.country("India");
				id.zipcode("411008");
				id.city("Pune");
				id.occupation("Employee");
				id.hobbies(true,false,false,true,false);
				id.submitInsurantData();
				//Enter Product Data
				Calendar.getInstance();
				cal.add(Calendar.DATE, 33);
				SimpleDateFormat s1 = new SimpleDateFormat("MM/dd/yyyy");
				String futureDate = s1.format(new Date(cal.getTimeInMillis()));
				pd.startdate(futureDate);
				pd.insurancesum(1);
				pd.damageinsurance(1);
				pd.orignalProducts();
				pd.submitProductData();
				cp.checkPrice(price, claim, dis, cover, type,testname);		
				sq.email("nikitaokeshwani@gmail.com");
				sq.username("Nikita");
				sq.password("Batra@123");
				sq.confirmpassword("Batra@123");
				sq.comment();
				sq.sendemail();	
				sq.message();
	}
	@Test(priority=3,dataProvider = "motorcycleData")
	public void motorcycleTest(String price, String claim, String dis, String cover, String type,String testname) throws InterruptedException, IOException {
		driver.findElement(By.partialLinkText("Motor")).click();
		//Enter Vehicle Data(Motorcycle)
				vd.selectMake("Honda");
				vd.modelMotorcycle(4);
				vd.cylindercapacity("250");
				vd.setEnginePerformance("1000");
				Calendar cal = Calendar.getInstance();
				cal.add(Calendar.DATE, -1);
				SimpleDateFormat s = new SimpleDateFormat("MM/dd/yyyy");
				String manDate = s.format(new Date(cal.getTimeInMillis()));
				vd.setManDate(manDate);
				vd.seatsmotorcycle("2");
			    vd.listprice("7000");			
				vd.annualmileage("5000");
				vd.submitVehicleData();
				//Enter Insurant data
				id.firstname("Nikita");
				id.lastname("Batra");
				id.birthdate("04/20/1992");
				id.gender();
				id.streetaddress("Bavdhan");
				id.country("India");
				id.zipcode("411008");
				id.city("Pune");
				id.occupation("Employee");
				id.hobbies(true,false,false,true,false);
				id.submitInsurantData();
				//Enter Product Data
				Calendar.getInstance();
				cal.add(Calendar.DATE, 33);
				SimpleDateFormat s1 = new SimpleDateFormat("MM/dd/yyyy");
				String futureDate = s1.format(new Date(cal.getTimeInMillis()));
				pd.startdate(futureDate);
				pd.insurancesum(1);
				pd.damageinsurance(1);
				pd.orignalProducts();
				pd.submitProductData();
				cp.checkPrice(price, claim, dis, cover, type,testname);		
				sq.email("nikitaokeshwani@gmail.com");
				sq.username("Nikita");
				sq.password("Batra@123");
				sq.confirmpassword("Batra@123");
				sq.comment();
				sq.sendemail();	
				sq.message();
	}
	@Test(priority=4,dataProvider = "camperData")
	public void camperTest(String price, String claim, String dis, String cover, String type,String testname) throws InterruptedException, IOException {
		driver.findElement(By.partialLinkText("Camper")).click();
		//Enter Vehicle Data(Truck)
				vd.selectMake("Toyota");
				vd.setEnginePerformance("1000");
				Calendar cal = Calendar.getInstance();
				cal.add(Calendar.DATE, -1);
				SimpleDateFormat s = new SimpleDateFormat("MM/dd/yyyy");
				String manDate = s.format(new Date(cal.getTimeInMillis()));
				vd.setManDate(manDate);
				vd.seats("9");
				vd.fuel("Petrol");
				vd.payload("900");
				vd.totalweight("50000");
				vd.listprice("90000");
				vd.annualmileage("5000");
				vd.submitVehicleData();
				//Enter Insurant data
				id.firstname("Nikita");
				id.lastname("Batra");
				id.birthdate("04/20/1992");
				id.gender();
				
				id.country("India");
				id.zipcode("411008");
				id.city("Pune");
				id.occupation("Employee");
				id.hobbies(true,false,true,false,false);
				id.submitInsurantData();
				//Enter Product Data
				Calendar.getInstance();
				cal.add(Calendar.DATE, 33);
				SimpleDateFormat s1 = new SimpleDateFormat("MM/dd/yyyy");
				String futureDate = s1.format(new Date(cal.getTimeInMillis()));
				pd.startdate(futureDate);
				pd.insurancesum(1);
				pd.damageinsurance(1);
				pd.orignalProducts();
				pd.submitProductData();
				cp.checkPrice(price, claim, dis, cover, type,testname);		
				sq.email("harsha@gmail.com");
				sq.username("Harsha");
				sq.password("Nanwani@123");
				sq.confirmpassword("Nanwani@123");
				sq.comment();
				sq.sendemail();	
				sq.message();
	}
	
	@DataProvider
	public Object[][] automobileData() {
		return new Object[][] {
			new Object[] { "274.00", "No", "No", "No", "Silver" ,"automobile"},
			new Object[] { "809.00", "Submit", "2", "Limited", "Gold","Null" },
			new Object[] { "1,589.00", "Submit", "5", "Limited", "Platinum","Null" },
			new Object[] { "3,026.00", "Submit", "10", "Unlimited", "Ultimate" ,"Null"},
		};
	}
	@DataProvider
	public Object[][] truckData() {
		return new Object[][] {
			new Object[] { "332.00", "No", "No", "No", "Silver" ,"truck"},
			new Object[] { "979.00", "Submit", "2", "Limited", "Gold" ,"Null"},
			new Object[] { "1,921.00", "Submit", "5", "Limited", "Platinum","Null" },
			new Object[] { "3,660.00", "Submit", "10", "Unlimited", "Ultimate","Null" },
		};
	}
	@DataProvider
	public Object[][] motorcycleData() {
		return new Object[][] {
			new Object[] { "251.00", "No", "No", "No", "Silver","motorcycle"},
			new Object[] { "741.00", "Submit", "2", "Limited", "Gold" ,"Null"},
			new Object[] { "1,455.00", "Submit", "5", "Limited", "Platinum","Null" },
			new Object[] { "2,772.00", "Submit", "10", "Unlimited", "Ultimate" ,"Null"},
		};
	}
		@DataProvider
		public Object[][] camperData() {
			return new Object[][] {
				new Object[] { "360.00", "No", "No", "No", "Silver" ,"camper"},
				new Object[] { "1,061.00", "Submit", "2", "Limited", "Gold" ,"Null"},
				new Object[] { "2,082.00", "Submit", "5", "Limited", "Platinum" ,"Null"},
				new Object[] { "3,967.00", "Submit", "10", "Unlimited", "Ultimate" ,"Null"},
			};
	}
	@BeforeTest(alwaysRun=true)
	public void beforeTest() {
		driver = new ChromeDriver();
		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(10));
		driver.get("https://sampleapp.tricentis.com/101/index.php");

		vd = new T01EnterVehicleData(driver);
		id= new T02EnterInsurantData(driver);
		pd= new T03EnterProductData(driver);
		cp= new T04SelectPriceOption(driver);
		sq= new T05SendQuote(driver);
		
	}

	@AfterTest
	public void afterTest() {
		driver.close();
	}

}
