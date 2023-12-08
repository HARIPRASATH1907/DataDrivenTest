package registrationTest;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.Select;

public class RegistrationTest {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		System.setProperty("webdriver.chrome.driver","C:\\SELENIUM\\Driver\\chromedriver.exe");
		ChromeOptions options=new ChromeOptions();
		options.addArguments("--remote-allow-origins=*");
		options.addArguments("start-maximized");
		WebDriver driver=new ChromeDriver(options);
	    driver.get("http://newtours.demoaut.com/");	
		FileInputStream file=new FileInputStream("C:\\Selenium\\Registration.xlsx");
		XSSFWorkbook workbook=new XSSFWorkbook(file);
		
		XSSFSheet sheet=workbook.getSheet("Sheet1");  //Providing Sheet Name
		
		// XSSFSheet sheet=workbook.getSheetAt(0);  //Providing Sheet Name

		int noofRows=sheet.getLastRowNum();          //Returns the row count
		
		System.out.println("No. of Records in the Excel Sheet:" + noofRows);
		
		for(int row=1;row<=noofRows;row++)
		{
			XSSFRow current_row=sheet.getRow(row);
			
			String First_Name=current_row.getCell(0).getStringCellValue();
			String Last_Name=current_row.getCell(1).getStringCellValue();
			String Phone=current_row.getCell(2).getStringCellValue();
			String Email=current_row.getCell(3).getStringCellValue();
			String Address=current_row.getCell(4).getStringCellValue();
			String City=current_row.getCell(5).getStringCellValue();
			String State=current_row.getCell(6).getStringCellValue();
			String PinCode=current_row.getCell(7).getStringCellValue();
			String Country=current_row.getCell(8).getStringCellValue();
			String UserName=current_row.getCell(9).getStringCellValue();
			String Password=current_row.getCell(10).getStringCellValue();
		
		//REgistration Process
			driver.findElement(By.linkText("REGISTER")).click();
			
			//Entering Contact Information
			
			driver.findElement(By.name("firstName")).sendKeys(First_Name);
			driver.findElement(By.name("lastName")).sendKeys(Last_Name);
			driver.findElement(By.name("phone")).sendKeys(Phone);
			driver.findElement(By.name("userName")).sendKeys(Email);
		
			//Entering Mailing Information
			
			driver.findElement(By.name("address1")).sendKeys(Address);
			driver.findElement(By.name("address2")).sendKeys(Address);
			driver.findElement(By.name("city")).sendKeys(City);
			driver.findElement(By.name("state")).sendKeys(State);
			driver.findElement(By.name("postalCode")).sendKeys(PinCode);
			
			//Country Selection
			Select dropcountry=new Select(driver.findElement(By.name("country")));
			dropcountry.selectByVisibleText(Country);
			
			// Entering User Information
			
			driver.findElement(By.name("email")).sendKeys(UserName);
			driver.findElement(By.name("password")).sendKeys(Password);
			driver.findElement(By.name("confirmPassword")).sendKeys(Password);
			
			//Submitting Form
			
			driver.findElement(By.name("register")).click();
			
			if(driver.getPageSource().contains("Thank you for registering")) {
				System.out.println("Registeration Completed for " + row + " record");
				
			}
			
			else
			{
				System.out.println("Registeration Failed for " + row + " record");
			}
			
			
		}
		
		System.out.println("Data Driven Test Completed");
		driver.close();	
			driver.quit();
			
			file.close();
					
	}

}
