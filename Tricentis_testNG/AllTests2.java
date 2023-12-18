package project;

import org.testng.annotations.Test;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.Calendar;
import java.util.Date;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.DataProvider;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.AfterTest;

public class AllTests2 {
	WebDriver driver;
	T01EnterVehicleData vd;
	T02EnterInsurantData id;
	T03EnterProductData pd;
	PriceAlltest2 cp;
	T05SendQuote sq;
	String fpath = "D:\\tricentisnew.xlsx";
	File file;
	FileInputStream fis;
	XSSFWorkbook wb;
	XSSFSheet sheet;
	XSSFRow row;
	XSSFCell cell;
	int index1 = 1;
    int rows,cells;
	@Test(priority = 1, dataProvider = "automobileData")
	public void automobileTest(String price, String claim, String dis, String cover, String type, String testname)
			throws InterruptedException, IOException {
		driver.findElement(By.partialLinkText("Auto")).click();

		// Enter Vehicle Data(Automobile)
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

		// Enter Insurant data
		id.firstname("Nikita");
		id.lastname("Batra");
		id.birthdate("04/20/1992");
		id.gender();
		id.streetaddress("Bavdhan");
		id.country("India");
		id.zipcode("411008");
		id.city("Pune");
		id.occupation("Employee");
		id.hobbies(true, true, false, false, false);
		id.submitInsurantData();

		// Enter Product Data
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
		testname = "Automobile";
		cp.checkPrice(price, claim, dis, cover, type, testname);
		sq.email("nikitaokeshwani@gmail.com");
		sq.username("Nikita");
		sq.password("Batra@123");
		sq.confirmpassword("Batra@123");
		sq.comment();
		sq.sendemail();
		sq.message();
	}

	@Test(priority = 2, dataProvider = "truckData")
	public void truckTest(String price, String claim, String dis, String cover, String type, String testname)
			throws InterruptedException, IOException {
		driver.findElement(By.partialLinkText("Truck")).click();

		// Enter Vehicle Data(Truck)
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
		// Enter Insurant data
		id.firstname("Nikita");
		id.lastname("Batra");
		id.birthdate("04/20/1992");
		id.gender();
		id.streetaddress("Bavdhan");
		id.country("India");
		id.zipcode("411008");
		id.city("Pune");
		id.occupation("Employee");
		id.hobbies(true, false, false, true, false);
		id.submitInsurantData();
		// Enter Product Data
		Calendar.getInstance();
		cal.add(Calendar.DATE, 33);
		SimpleDateFormat s1 = new SimpleDateFormat("MM/dd/yyyy");
		String futureDate = s1.format(new Date(cal.getTimeInMillis()));
		pd.startdate(futureDate);
		pd.insurancesum(1);
		pd.damageinsurance(1);
		pd.orignalProducts();
		pd.submitProductData();
		testname = "Truck";
		cp.checkPrice(price, claim, dis, cover, type, testname);
		sq.email("nikitaokeshwani@gmail.com");
		sq.username("Nikita");
		sq.password("Batra@123");
		sq.confirmpassword("Batra@123");
		sq.comment();
		sq.sendemail();
		sq.message();
	}

	@Test(priority = 3, dataProvider = "motorData")
	public void motorcycleTest(String price, String claim, String dis, String cover, String type, String testname)
			throws InterruptedException, IOException {
		driver.findElement(By.partialLinkText("Motor")).click();
		// Enter Vehicle Data(Motorcycle)
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
		// Enter Insurant data
		id.firstname("Nikita");
		id.lastname("Batra");
		id.birthdate("04/20/1992");
		id.gender();
		id.streetaddress("Bavdhan");
		id.country("India");
		id.zipcode("411008");
		id.city("Pune");
		id.occupation("Employee");
		id.hobbies(true, false, false, true, false);
		id.submitInsurantData();
		// Enter Product Data
		Calendar.getInstance();
		cal.add(Calendar.DATE, 33);
		SimpleDateFormat s1 = new SimpleDateFormat("MM/dd/yyyy");
		String futureDate = s1.format(new Date(cal.getTimeInMillis()));
		pd.startdate(futureDate);
		pd.insurancesum(1);
		pd.damageinsurance(1);
		pd.orignalProducts();
		pd.submitProductData();
		testname = "Motorcycle";
		cp.checkPrice(price, claim, dis, cover, type, testname);
		sq.email("nikitaokeshwani@gmail.com");
		sq.username("Nikita");
		sq.password("Batra@123");
		sq.confirmpassword("Batra@123");
		sq.comment();
		sq.sendemail();
		sq.message();
	}

	@Test(priority = 4, dataProvider = "camperData")
	public void camperTest(String price, String claim, String dis, String cover, String type, String testname)
			throws InterruptedException, IOException {
		driver.findElement(By.partialLinkText("Camper")).click();

		// Enter Vehicle Data(Truck)
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
		// Enter Insurant data
		id.firstname("Nikita");
		id.lastname("Batra");
		id.birthdate("04/20/1992");
		id.gender();

		id.country("India");
		id.zipcode("411008");
		id.city("Pune");
		id.occupation("Employee");
		id.hobbies(true, false, true, false, false);
		id.submitInsurantData();
		// Enter Product Data
		Calendar.getInstance();
		cal.add(Calendar.DATE, 33);
		SimpleDateFormat s1 = new SimpleDateFormat("MM/dd/yyyy");
		String futureDate = s1.format(new Date(cal.getTimeInMillis()));
		pd.startdate(futureDate);
		pd.insurancesum(1);
		pd.damageinsurance(1);
		pd.orignalProducts();
		pd.submitProductData();
		testname = "Camper";
		cp.checkPrice(price, claim, dis, cover, type, testname);
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
		sheet = wb.getSheetAt(0);
		row=sheet.getRow(1);
		rows = sheet.getPhysicalNumberOfRows();
		cells= sheet.getRow(1).getPhysicalNumberOfCells();
		String[][] data = new String[rows-1][6];
		for (int i = 0; i < rows - 1; i++) {
			row = sheet.getRow(i + 1);
			if (row != null) {
				for (int j = 0; j < cells; j++) {
					cell = row.getCell(j);
					if (cell != null) {
						if (cell.getCellType() == CellType.NUMERIC) {
							data[i][j] = String.valueOf((int) cell.getNumericCellValue());
						} else {
							data[i][j] = cell.getStringCellValue();
						}
					} else {
						// Handle null cell
						data[i][j] = ""; // Or any other suitable value
					}
				}
			}
		}
		return data;
	}

	@DataProvider
	public Object[][] truckData() {
		sheet = wb.getSheetAt(1);
		row=sheet.getRow(1);
		rows = sheet.getPhysicalNumberOfRows();
		cells= sheet.getRow(1).getPhysicalNumberOfCells();
		String[][] data = new String[rows-1][6];
		for (int i = 0; i < rows - 1; i++) {
			row = sheet.getRow(i + 1);
			if (row != null) {
				for (int j = 0; j < cells; j++) {
					cell = row.getCell(j);
					if (cell != null) {
						if (cell.getCellType() == CellType.NUMERIC) {
							data[i][j] = String.valueOf((int) cell.getNumericCellValue());
						} else {
							data[i][j] = cell.getStringCellValue();
						}
					} else {
						// Handle null cell
						data[i][j] = ""; // Or any other suitable value
					}
				}
			}
		}
		return data;
	}

	@DataProvider
	public Object[][] motorData() {
		sheet = wb.getSheetAt(2);
		row=sheet.getRow(1);
		 rows = sheet.getPhysicalNumberOfRows();
		cells= sheet.getRow(1).getPhysicalNumberOfCells();
		String[][] data = new String[rows-1][6];
		for (int i = 0; i < rows - 1; i++) {
			row = sheet.getRow(i + 1);
			if (row != null) {
				for (int j = 0; j < cells; j++) {
					cell = row.getCell(j);
					if (cell != null) {
						if (cell.getCellType() == CellType.NUMERIC) {
							data[i][j] = String.valueOf((int) cell.getNumericCellValue());
						} else {
							data[i][j] = cell.getStringCellValue();
						}
					} else {
						// Handle null cell
						data[i][j] = ""; // Or any other suitable value
					}
				}
			}
		}
		return data;
	}

	@DataProvider
	public Object[][] camperData() {
		sheet = wb.getSheetAt(3);
		row=sheet.getRow(1);
		rows = sheet.getPhysicalNumberOfRows();
		cells= sheet.getRow(1).getPhysicalNumberOfCells();
		String[][] data = new String[rows-1][6];
		for (int i = 0; i < rows - 1; i++) {
			row = sheet.getRow(i + 1);
			if (row != null) {
				for (int j = 0; j < cells; j++) {
					cell = row.getCell(j);
					if (cell != null) {
						if (cell.getCellType() == CellType.NUMERIC) {
							data[i][j] = String.valueOf((int) cell.getNumericCellValue());
						} else {
							data[i][j] = cell.getStringCellValue();
						}
					} else {
						// Handle null cell
						data[i][j] = ""; // Or any other suitable value
					}
				}
			}
		}
		return data;
	}

	@AfterMethod
	public void afterMethod() {

	}

	@BeforeTest
	public void beforeTest() throws IOException {
		driver = new ChromeDriver();
		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(10));
		driver.get("https://sampleapp.tricentis.com/101/index.php");
		vd = new T01EnterVehicleData(driver);
		id = new T02EnterInsurantData(driver);
		pd = new T03EnterProductData(driver);
		cp = new PriceAlltest2(driver);
		sq = new T05SendQuote(driver);

		file = new File(fpath);
		fis = new FileInputStream(file);
		wb = new XSSFWorkbook(fis);
		sheet = wb.getSheetAt(0);
		row = sheet.getRow(0);
	}

	@AfterTest
	public void afterTest() throws IOException {
		wb.close();
		fis.close();
		driver.close();
	}

}
