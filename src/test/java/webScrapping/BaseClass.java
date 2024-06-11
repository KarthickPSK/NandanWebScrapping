package webScrapping;

import java.io.File;
import java.io.FileInputStream;
import java.util.Properties;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.safari.SafariDriver;

public class BaseClass {
	
	public File file;
	public FileInputStream fis;
	public Properties prop;
	public WebDriver driver;
	public String xlPath="C:\\Backup\\Nandan Bro\\WebPage.xlsx";
	
	public BaseClass() {
		file=new File("./configuration\\config.properties");
		try {
		fis=new FileInputStream(file);
		prop=new Properties();
		prop.load(fis);
		}
		catch(Exception e) {
			e.printStackTrace();
		}
	}

	public WebDriver selectBrowser(String browserName) {
		if(browserName.equalsIgnoreCase("Chrome")) {
			ChromeOptions option=new ChromeOptions();
			option.addArguments("--headless");
			driver=new ChromeDriver(option);
		}else if(browserName.equalsIgnoreCase("Firefox")) {
			driver=new FirefoxDriver();
		}else if(browserName.equalsIgnoreCase("edge")) {
			driver=new EdgeDriver();
		}else if(browserName.equalsIgnoreCase("IE")) {
			driver = new InternetExplorerDriver();
		}else if(browserName.equalsIgnoreCase("Safari")) {
			driver=new SafariDriver();
		}else {
			driver=new ChromeDriver();
		}
		return driver;
	}
}
