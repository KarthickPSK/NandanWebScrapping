package webScrapping;

import java.io.File;
import java.io.IOException;

import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import utilities.ExcelUtilities;

public class DataProvide {
	String path="C:\\Backup\\Nandan Bro\\NandanWebScrapping.xlsx";
	
	
//	@Test(dataProvider="excelData")
//	public void execute(String n1) {
//		System.out.println(n1);
//		
//	}
	
	@DataProvider(name="webScrappingData")
	public String[] readData() throws IOException {
		int rowCount = ExcelUtilities.getRowCount(path,"Sheet1");
		String data[]=new String[rowCount+1];
		for(int i=0;i<=rowCount;i++) {
				data[i]=ExcelUtilities.getCellData(path,"Sheet1",i,0);
		}
		return data;	
	}

}
