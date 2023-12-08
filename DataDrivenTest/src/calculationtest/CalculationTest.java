package calculationtest;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.Select;

public class CalculationTest {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub

		System.setProperty("webdriver.chrome.driver","C:\\SELENIUM\\Driver\\chromedriver.exe");
		ChromeOptions options=new ChromeOptions();
		options.addArguments("--remote-allow-origins=*");
		options.addArguments("start-maximized");
		WebDriver driver=new ChromeDriver(options);
	    driver.get("http://newtours.demoaut.com/");	
		FileInputStream file=new FileInputStream("C:\\Selenium\\CalData.xlsx");
		XSSFWorkbook workbook=new XSSFWorkbook(file);
		
		XSSFSheet sheet=workbook.getSheet("Sheet1");  //Providing Sheet Name
		

		int rowcount=sheet.getLastRowNum();          //Returns the row count
		
		
		for(int i=1;i<=rowcount;i++)
		{
			XSSFRow row=sheet.getRow(i);
			
			XSSFCell principlecell=row.getCell(0);
			int princ=(int)principlecell.getNumericCellValue();
			
			XSSFCell roi=row.getCell(0);
			int rateofinterest=(int)roi.getNumericCellValue();
			
			XSSFCell period=row.getCell(0);
			int per=(int)period.getNumericCellValue();
			
			XSSFCell Frequency=row.getCell(0);
			int Freq=(int)Frequency.getNumericCellValue();
			
			XSSFCell MaturityValue=row.getCell(0);
			int Exp_mvalue=(int)MaturityValue.getNumericCellValue();
			
			
			driver.findElement(By.id("principal")).sendKeys(String.valueOf(princ));
			driver.findElement(By.id("interest")).sendKeys(String.valueOf(rateofinterest));
			driver.findElement(By.id("tenure")).sendKeys(String.valueOf(per));
			
			
			Select periodcombo=new Select(driver.findElement(By.id("tenurePeriod")));
			periodcombo.selectByVisibleText("year(s)");
			
			Select frequency=new Select(driver.findElement(By.id("frequency")));
			frequency.selectByVisibleText("Freq");
			
			driver.findElement(By.xpath(".//*[@id='fdMatVal']/div[2]/a[1]/img")).click();
			
		String actual_mvalue=	driver.findElement(By.xpath(".//*[@id='resp_matval']/strong")).getText();
		
		if((Double.parseDouble(actual_mvalue)==Exp_mvalue))
		{
			System.out.println("Test Passed");
		}
		
		else {
			System.out.println("Test Failed");
		}
		
		
		driver.findElement(By.xpath(".//*[@id='fdMatVal']/div[2]/a[2]/img")).click();
		
		
	}
driver.close();
driver.quit();

}
}