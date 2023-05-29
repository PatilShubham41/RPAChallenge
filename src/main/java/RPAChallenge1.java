import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

public class RPAChallenge1 {
	
	public  static String getCellValue(String colHeader, int rowNumber, Sheet sheet) {

		XSSFRow row = (XSSFRow) sheet.getRow(0);
		XSSFRow row1 = (XSSFRow) sheet.getRow(rowNumber - 1);

		for (int j = 0; j < row.getLastCellNum(); j++) {

			if (row.getCell(j).getStringCellValue().equals(colHeader))
				if (row1.getCell(j).getCellType() == CellType.NUMERIC)
					return String.valueOf((long)(row1.getCell(j).getNumericCellValue()));
				else
					return row1.getCell(j).getStringCellValue();
		}
		
		System.out.println("Some problem with excel");
		
		return null;
	}

	public static void main(String[] args) {

		System.setProperty("webdriver.chrome.driver", "C:\\dev\\Java\\Selenium\\drivers\\chromedriver.exe");
		WebDriver driver = new ChromeDriver();
		
		driver.manage().window().maximize();
		driver.manage().timeouts().pageLoadTimeout(60, TimeUnit.SECONDS);
		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		driver.manage().deleteAllCookies();
		driver.get("https://www.rpachallenge.com/");
		
		Sheet sheet = null;
		//Read Excel
		try {
			File file = new File("C:\\Users\\Shubham\\Downloads\\challenge.xlsx");
			FileInputStream inputStream = new FileInputStream(file);
			Workbook testdata = new XSSFWorkbook(inputStream);
			sheet = testdata.getSheet("Sheet1");

		} catch (IOException e) {
			e.printStackTrace();
		}
		
		driver.findElement(By.xpath("//button[text()='Start']")).click();
		for(int i = 2; i < 12; i++)
		{
			driver.findElement(By.cssSelector("input[ng-reflect-name='labelCompanyName']")).sendKeys(getCellValue("Company Name", i, sheet));
			driver.findElement(By.cssSelector("input[ng-reflect-name='labelEmail']")).sendKeys(getCellValue("Email", i, sheet));
			driver.findElement(By.cssSelector("input[ng-reflect-name='labelPhone']")).sendKeys(getCellValue("Phone Number", i, sheet));
			driver.findElement(By.cssSelector("input[ng-reflect-name='labelAddress']")).sendKeys(getCellValue("Address", i, sheet));
			driver.findElement(By.cssSelector("input[ng-reflect-name='labelRole']")).sendKeys(getCellValue("Role in Company", i, sheet));
			
	
			driver.findElement(By.cssSelector("input[ng-reflect-name='labelFirstName']")).sendKeys(getCellValue("First Name", i, sheet));
			driver.findElement(By.cssSelector("input[ng-reflect-name='labelLastName']")).sendKeys(getCellValue("Last Name", i, sheet));
			
			
			
			driver.findElement(By.cssSelector("input[type=\"submit\"][value=\"Submit\"]")).click();
			
		}
		
		System.out.println(driver.findElement(By.cssSelector("div.message2")).getText());
		
		driver.close();

	}

}
