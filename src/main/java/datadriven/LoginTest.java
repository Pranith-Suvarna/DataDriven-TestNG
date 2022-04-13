package datadriven;

import java.io.IOException;
import java.util.concurrent.TimeUnit;
import java.util.function.Function;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;
import org.testng.ITestResult;
import org.testng.annotations.AfterClass;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import io.github.bonigarcia.wdm.WebDriverManager;

public class LoginTest {
	WebDriver driver;
	static int count = 1 ;

	@BeforeClass
	public void setup()
	{
		WebDriverManager.chromedriver().setup();
	  //System.setProperty("webdriver.chrome.driver","./drivers/chromedriver.exe");
		driver=new ChromeDriver();
		driver.manage().timeouts().implicitlyWait(5,TimeUnit.SECONDS);
		driver.manage().window().maximize();
	}
	
	
	@Test(dataProvider="LoginData")
	public void loginTest(String user,String pwd,String expected,String status)
	{
		driver.get("https://admin-demo.nopcommerce.com/login");
		
		WebElement txtEmail=driver.findElement(By.id("Email"));
		txtEmail.clear();
		txtEmail.sendKeys(user);
		
		
		WebElement txtPassword=driver.findElement(By.id("Password"));
		txtPassword.clear();
		txtPassword.sendKeys(pwd);
		
		driver.findElement(By.xpath("//*[@class='button-1 login-button']")).click(); //Login  button
		
		String exp_title="Dashboard / nopCommerce administration";
		String act_title=driver.getTitle();
		
		if(expected.equals("Valid"))
		{
			if(exp_title.equals(act_title))
			{
				//driver.findElement(By.linkText("Logout")).click();
				(new WebDriverWait(driver, 60L)).until((Function)ExpectedConditions.elementToBeClickable(By.xpath("//a[text()='Logout']")));
				driver.findElement(By.xpath("//a[text()='Logout']")).click();
				Assert.assertTrue(true);
			}
			else
			{
				Assert.assertTrue(false);
			}
		}
		else if(expected.equals("Invalid"))
		{
			if(exp_title.equals(act_title))
			{
				//driver.findElement(By.linkText("Logout")).click();
				(new WebDriverWait(driver, 60L)).until((Function)ExpectedConditions.elementToBeClickable(By.xpath("//a[text()='Logout']")));
				driver.findElement(By.xpath("//a[text()='Logout']")).click();

				Assert.assertTrue(false);
			}
			else
			{
				Assert.assertTrue(true);
			}
		}
		
	}
	
	@DataProvider(name="LoginData")
	public String [][] getData() throws IOException
	{
		/*String loginData[][]= {
								{"admin@yourstore.com","admin","Valid"},
								{"admin@yourstore.com","adm","Invalid"},
								{"adm@yourstore.com","admin","Invalid"},
								{"adm@yourstore.com","adm","Invalid"}
							};*/
		
		//get the data from excel
		String path=".\\datafiles\\loginData.xlsx";
		XLUtility xlutil=new XLUtility(path);
		
		int totalrows=xlutil.getRowCount("Sheet1");
		int totalcols=xlutil.getCellCount("Sheet1",1);	

				
		String loginData[][]=new String[totalrows][totalcols];
			
		
		for(int i=1;i<=totalrows;i++) //1
		{
			for(int j=0;j<totalcols;j++) //0
			{
				loginData[i-1][j]=xlutil.getCellData("Sheet1", i, j);
				
			}
				
		}
		
		return loginData;
	}
	
	@AfterMethod
	public void writeResult(ITestResult result) throws IOException
	{
		System.out.println("method name:" + result.getMethod().getMethodName());

		String path = ".\\datafiles\\loginData.xlsx";
		XLUtility xlutil = new XLUtility(path);
		
		xlutil.updateResult(path, "Sheet1" ,result);
		count++;
	}
	
	@AfterClass
	void tearDown()
	{
		driver.quit();
	}
	
}
