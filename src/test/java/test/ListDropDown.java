package test;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.jar.Attributes.Name;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;

import io.github.bonigarcia.wdm.WebDriverManager;

public class ListDropDown {

	public static void main(String[] args) throws IOException {
		//System.setProperty("webdriver.chrome.driver", "D:\\selenium\\chromedriver.exe");
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet0 = workbook.createSheet("data1");
		File file = new File("C:\\Users\\koira\\eclipse-workspace\\ExcelDriven\\excel\\dropdown.xlsx");
		FileOutputStream fo = new FileOutputStream(file);
		workbook.write(fo);
		
		
		
		WebDriverManager.chromedriver().setup();
		WebDriver driver = new ChromeDriver();
		driver.navigate().to("http://newtours.demoaut.com/mercuryregister.php");

		WebElement country = driver.findElement(By.name("country"));

		Select sel = new Select(country);
		List<WebElement> list = sel.getOptions();
		System.out.println("Number of country " + list.size());
		 sel.selectByVisibleText("ALBANIA ");
		
		
		
		
		
		for(int rows=0;rows<10;rows++)
		{
			XSSFRow row = sheet0.createRow(rows);
			
			for(int cols=0;cols<10;cols++)
			{
				Cell cell = row.createCell(cols);
			   for (WebElement allList : list)
			   {
				   cell.setCellValue(allList.getText());
			   }
			}
		}
	
		fo.close();
		System.out.println("File is created");
			 
	}

}
