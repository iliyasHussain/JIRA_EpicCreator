package epicAutomator;
import java.io.File;
import java.io.FileInputStream;
import java.util.concurrent.TimeUnit;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.apache.poi.xssf.usermodel.XSSFSheet;

public class EpicAutomator {
	public static void main(String[] args) throws Exception {
		
		//fetch raw data from excel 
		File src = new File("Enter FilePath here");
		FileInputStream fis = new FileInputStream(src);
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		XSSFSheet sh = wb.getSheetAt(0);
		int Rownum = sh.getLastRowNum();
		System.out.println("Total No. of Apps = " + Rownum);

		//setup webDriver and login to JIRA//
		System.setProperty("webdriver.chrome.driver", "C:/ChromeDriver/chromedriver.exe");
		WebDriver driver = new ChromeDriver();
		driver.manage().window().maximize();
		driver.get("https://issues.labcollab.net/login.jsp");
		driver.manage().timeouts().implicitlyWait(120, TimeUnit.SECONDS);
		WebElement element = driver.findElement(By.id("login-form-username"));
		element.sendKeys("Enter UserName here");
		WebElement element1 = driver.findElement(By.id("login-form-password"));
		element1.sendKeys("Enter Password here");
		element.submit();
		driver.manage().timeouts().implicitlyWait(120, TimeUnit.SECONDS);
        
		//Read and Store data from excel file//
		for(int i=1; i<=Rownum; i++) {
			String AppName[] = new String[i+1];
			String ASIN[] = new String[i+1];
			String Data[] = new String[i+1];

			AppName[i]= sh.getRow(i).getCell(0).getStringCellValue();
			ASIN[i]= sh.getRow(i).getCell(1).getStringCellValue();
			Data[i] = AppName[i] + " " + "(" + ASIN[i] +")";
			System.out.println("App Name and ASIN are: " + Data[i]);
            //End Of Read and Store data from excel file module//

            // Click on "Create" button to create an issue  
			WebDriverWait wait = new WebDriverWait(driver, 120);
			WebElement element2 = driver.findElement(By.id("create_link"));
			wait.until(ExpectedConditions.elementToBeClickable(element2)).click();

			// Select issue type
			driver.manage().timeouts().implicitlyWait(120, TimeUnit.SECONDS);
			WebDriverWait waitForIt = new WebDriverWait(driver, 30);
			WebElement elem = waitForIt.until(ExpectedConditions.elementToBeClickable(By.cssSelector("#issuetype-field")));
			elem.click();
			elem.sendKeys("Epic");
			elem.sendKeys(Keys.RETURN);
			
			// Set Epic name
			driver.manage().timeouts().implicitlyWait(120, TimeUnit.SECONDS);
			elem = waitForIt.until(ExpectedConditions.elementToBeClickable(By.cssSelector("#customfield_10601")));
			elem.sendKeys(Data[i]);

			// Enter summary
			driver.manage().timeouts().implicitlyWait(120, TimeUnit.SECONDS);
			elem = waitForIt.until(ExpectedConditions.elementToBeClickable(By.cssSelector("#summary")));
			elem.clear();
			elem.sendKeys(Data[i]+" - Master JIRA for Compat Issues");

			// Set assignee
			driver.manage().timeouts().implicitlyWait(120, TimeUnit.SECONDS);
			elem = waitForIt.until(ExpectedConditions.elementToBeClickable(By.cssSelector("#assignee-field")));
			elem.clear();
			elem.sendKeys("Enter Assignee alias here");

			// Set component(s)
			driver.manage().timeouts().implicitlyWait(120, TimeUnit.SECONDS);
			elem = waitForIt.until(ExpectedConditions.elementToBeClickable(By.cssSelector("#components-textarea")));
			elem.clear();
			elem.sendKeys("Untriaged");
			elem.sendKeys(Keys.RETURN);

			// Set labels
			driver.manage().timeouts().implicitlyWait(120, TimeUnit.SECONDS);
			elem = waitForIt.until(ExpectedConditions.elementToBeClickable(By.cssSelector("#labels-textarea")));
			elem.clear();
			elem.sendKeys("Enter labels here");
			elem.sendKeys(Keys.RETURN);
			
			// Uncomment this to proceed with issue creation by clicking on the "Create" button
			//driver.manage().timeouts().implicitlyWait(120, TimeUnit.SECONDS);
			//elem = waitForIt.until(ExpectedConditions.elementToBeClickable(By.cssSelector("#create-issue-submit")));
			//elem.click();
			
			// Cancels issue creation. Comment this section out when you're sure the code works as intended. 
			elem.sendKeys(Keys.ESCAPE);
			Thread.sleep(1000);
			driver.switchTo().alert().accept();
			Thread.sleep(500);
            //End Of locate elements and pass value module//
		}
		wb.close();
		driver.close();
	}
}