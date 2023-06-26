package Base;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.remote.DesiredCapabilities;

import Utility.CommonFile;
import io.github.bonigarcia.wdm.WebDriverManager;

public class Browser extends CommonFile{
	
//Job Closing Journal	
	protected static WebDriver driver;

	public void initilization() throws Exception {
		System.setProperty("webdriver.http.factory", "jdk-http-client");
		WebDriverManager.chromedriver().setup();
		ChromeOptions options = new ChromeOptions();
		options.addArguments("--remote-allow-origins=*");
		DesiredCapabilities cp = new DesiredCapabilities();
		cp.setCapability(ChromeOptions.CAPABILITY, options);
		options.merge(cp);
		driver = new ChromeDriver(options);

		driver.get(readExcelFileFinal(4, 1));
		driver.manage().window().maximize();
	}
	

}
