package utilis;

import java.time.Duration;
import java.util.concurrent.TimeUnit;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.testng.annotations.AfterSuite;
import org.testng.annotations.BeforeSuite;

public class AppUtils 
{

	public static WebDriver driver;
	public static String url = "http://orangehrm.qedgetech.com";


	@BeforeSuite
	public  static void launchApp()
	{
		System.setProperty("webdriver.chrome.driver", "F://ChromeDriver//chromedriver.exe");
		ChromeOptions co=new ChromeOptions();
		co.addArguments("--remote-allow-origins=*");
		
		driver = new ChromeDriver(co);
		
		
		driver.manage().deleteAllCookies();
		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(10,TimeUnit.SECONDS);
		driver.get(url);		
	}

	@AfterSuite
	public static void closeApp()
	{
		driver.close();		
	}


}



