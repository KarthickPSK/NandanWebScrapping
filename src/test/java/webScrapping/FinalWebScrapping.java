package webScrapping;

import java.io.IOException;
import java.time.Duration;
import java.util.List;

import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.testng.annotations.Test;

import utilities.ExcelUtilities;



public class FinalWebScrapping extends BaseClass {
       String path="C:\\Backup\\Nandan Bro\\WebScrapping(3).xlsx";
       int v=3;
       int start=v;
       int nexttablestart=v;
       int titlefirst=v;
       int titlelast;
       int pertitle=v;
       String title;
      
    @Test(dataProvider="webScrappingData",dataProviderClass = DataProvide.class)   
	public void init(String data) throws IOException, InterruptedException {	
		ChromeOptions option=new ChromeOptions();
		option.addArguments("--headless");
		driver=new ChromeDriver(option);
		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(20));
		driver.get(ExcelUtilities.getCellData(xlPath,"Sheet1",0,1)+data+ExcelUtilities.getCellData(xlPath,"Sheet1",1,1));
		List<WebElement> tables=driver.findElements(By.xpath("//*[@class='shadowWrap-vKM0WfUu']/div[1]"));
		List<WebElement> titles = driver.findElements(By.xpath("//div[@class='title-OU0c3RqN']"));

		for(int t=0;t<tables.size();t++) {
			WebElement tab = tables.get(t);
			title=titles.get(t).getText();
		    setData(tab);
		    ExcelUtilities.setCellData(path,"Sheet1",pertitle,2,"Quarterly");
		    ExcelUtilities.setMergingCell(path,"Sheet1",pertitle,v,2,2);
		    v+=2;
		    pertitle=v;
		    int z=t+1;
		    driver.findElement(By.xpath("(//*[@id='FY'])["+z+"]")).click();
		    Thread.sleep(2000);
		    setData(tab);
		    ExcelUtilities.setCellData(path,"Sheet1",titlefirst,1,title);
		    ExcelUtilities.setMergingCell(path,"Sheet1",titlefirst,titlelast,1,1);
		    ExcelUtilities.setCellData(path,"Sheet1",pertitle,2,"Annual");
		    ExcelUtilities.setMergingCell(path,"Sheet1",pertitle,v,2,2);
		    v+=2;
			titlefirst=v;
			pertitle=v;
		}
		v=ExcelUtilities.getRowCount(path,"Sheet1");
//		System.out.println("One Company last row num: "+v);
//		System.out.println("Merge Start: "+start);
//		System.out.println("Merge End: "+v);
		ExcelUtilities.setCellData(path,"Sheet1", start,0, data);
//		System.out.println("Printed Data...");
		ExcelUtilities.setMergingCell(path,"Sheet1", start, v,0,0);
//		System.out.println("Merged Rows...");
		v+=4;
        pertitle=titlefirst=start=v;
	}

    
	public void setData(WebElement tab) throws IOException {

			List<WebElement> rows = tab.findElements(By.xpath("div"));
			for(int r=0;r<rows.size();r++) {
				WebElement row = rows.get(r);				
				int o=r+1;
				String text = tab.findElement(By.xpath("div["+o+"]/div[3]")).getText();
				ExcelUtilities.setCellData(path,"Sheet1",v,3, text);
				List<WebElement> datas = tab.findElements(By.xpath("div["+o+"]//div[contains(@class,'ner-OxVAcLqi')]"));
				for(int d=0;d<datas.size();d++) {
					String text2 = datas.get(d).getText();
					int y=d+4;
					ExcelUtilities.setCellData(path,"Sheet1",v,y, text2);
				}
				v++;
			}
			v=ExcelUtilities.getRowCount(path,"Sheet1");
//			System.out.println("Last Row num is "+v);
			titlelast=v;		
	}
	
	public void elementClick(WebElement element) {
		JavascriptExecutor js=(JavascriptExecutor) driver;
		js.executeScript("arguments[0].click();",element);
	}
	
	public void elementVisible(WebElement element) {
		JavascriptExecutor js=(JavascriptExecutor) driver;
		js.executeScript("arguments[0].scrollIntoView(true);",element);
	}
	
	
}