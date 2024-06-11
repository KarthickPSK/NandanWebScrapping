package utilities;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.io.File;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtilities {
	
	public static FileInputStream fis;
	public static FileOutputStream fos;
	public static XSSFWorkbook workbook;
	public static XSSFSheet sheet;
	public static XSSFRow row;
	public static XSSFCell cell;
	public static CellStyle style;
	
	
	public static int getRowCount(String path,String SheetName) throws IOException {
		fis=new FileInputStream(path);
		workbook=new XSSFWorkbook(fis);
		sheet=workbook.getSheet(SheetName);
		int rowcount=sheet.getLastRowNum();
		workbook.close();
		fis.close();
		return rowcount;
	}
	
	public static int getCellCount(String path, String SheetName,int rownum) throws IOException {
		fis=new FileInputStream(path);
		workbook=new XSSFWorkbook(fis);
		sheet=workbook.getSheet(SheetName);
		row=sheet.getRow(rownum);
		int cellcount=row.getLastCellNum();
		workbook.close();
		fis.close();
		return cellcount;
	}
	
	public static String getCellData(String path, String SheetName,int rownum,int colnum) throws IOException {
		fis=new FileInputStream(path);
		workbook=new XSSFWorkbook(fis);
		sheet=workbook.getSheet(SheetName);
		row=sheet.getRow(rownum);
		cell=row.getCell(colnum);
		DataFormatter formatter=new DataFormatter();
		String data;
		try {
			data=formatter.formatCellValue(cell);
		}
		catch(Exception e) {
			data=" ";
		}
		workbook.close();
		fis.close();
		return data;
	}

/*	public static void setCellData(String path,String SheetName,int rownum,int colnum,String data) throws IOException {
		fis=new FileInputStream(path);
		workbook=new XSSFWorkbook(fis);
		sheet=workbook.getSheet(SheetName);
		row=sheet.getRow(rownum);
		cell=row.createCell(colnum);
		cell.setCellValue(data);
		fos=new FileOutputStream(path);
		workbook.write(fos);
		workbook.close();
		fis.close();
		fos.close();
	}       */
	
	public static void setCellData(String path, String SheetName,int rownum,int colnum,String data) throws IOException {
		File xlfile=new File(path);
		if(!xlfile.exists()) {
	    workbook=new XSSFWorkbook();	
	    fos=new FileOutputStream(path);
	    workbook.write(fos);
		}
		fis=new FileInputStream(path);
		workbook=new XSSFWorkbook(fis);
		
		if(workbook.getSheetIndex(SheetName)==-1) {
			workbook.createSheet(SheetName);
		}
		sheet=workbook.getSheet(SheetName);
		
		if(sheet.getRow(rownum)==null) {
			sheet.createRow(rownum);
		}
		row=sheet.getRow(rownum);
		
		cell=row.createCell(colnum);
		cell.setCellValue(data);
		fos=new FileOutputStream(path);
		workbook.write(fos);
		workbook.close();
		fis.close();
		fos.close();
	}
	
	public static void fillGreenColor(String path, String SheetName,int rownum,int colnum) throws IOException {
		fis=new FileInputStream(path);
		workbook=new XSSFWorkbook(fis);
		sheet=workbook.getSheet(SheetName);
		row=sheet.getRow(rownum);
		cell=row.getCell(colnum);
		style=workbook.createCellStyle();
		style.setFillForegroundColor(IndexedColors.GREEN.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		cell.setCellStyle(style);
		workbook.write(fos);
		workbook.close();
		fis.close();
		fos.close();
	}
	
	public static void fillRedColor(String path, String SheetName,int rownum,int colnum) throws IOException {
		fis=new FileInputStream(path);
		workbook=new XSSFWorkbook(fis);
		sheet=workbook.getSheet(SheetName);
		row=sheet.getRow(rownum);
		cell=row.getCell(colnum);
		style=workbook.createCellStyle();
		style.setFillForegroundColor(IndexedColors.RED.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		cell.setCellStyle(style);
		workbook.write(fos);
		workbook.close();
		fis.close();
		fos.close();
	}
	
	public static void setMergingCell(String path,String sheetName,int firstRowNum,int lastRowNum,int firstColNum,int lastColNum) throws IOException {
		fis=new FileInputStream(path);
		workbook=new XSSFWorkbook(fis);
		sheet = workbook.getSheet(sheetName);
		sheet.addMergedRegion(new CellRangeAddress(firstRowNum,lastRowNum,firstColNum,lastColNum));
		fos=new FileOutputStream(path);
		workbook.write(fos);
		workbook.close();
		fos.close();
		fis.close();	
	}
	
	public static void setMergingCells(String path,String sheetName,String rowColName) throws IOException {
		fis=new FileInputStream(path);
		workbook=new XSSFWorkbook(fis);
		sheet=workbook.getSheet(sheetName);
		sheet.addMergedRegion(CellRangeAddress.valueOf(rowColName));
		fos=new FileOutputStream(path);
		workbook.write(fos);
		workbook.close();
		fos.close();
		fis.close();		
	}
	
	public static int getMergedCell(String path,String sheetName) throws IOException {
		fis=new FileInputStream(path);
		workbook=new XSSFWorkbook(fis);
		sheet=workbook.getSheet(sheetName);
		int numMergedRegions = sheet.getNumMergedRegions();
		workbook.close();
		fis.close();
		return numMergedRegions; 	
	}
	
	public static int getMergedCells(String path,String sheetName) throws IOException {
		fis=new FileInputStream(path);
		workbook=new XSSFWorkbook(fis);
		sheet=workbook.getSheet(sheetName);
		List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
		int size = mergedRegions.size();
		workbook.close();
		fis.close();
		return size;		
	}
	
	public static void setUnmergingCell(String path,String sheetName,int mergedNum) throws IOException {
		fis=new FileInputStream(path);
		workbook=new XSSFWorkbook(fis);
		sheet=workbook.getSheet(sheetName);
		sheet.removeMergedRegion(mergedNum);
		fos=new FileOutputStream(path);
		workbook.write(fos);
		workbook.close();
		fos.close();
		fis.close();
	}
	
    public static void mergeAndSetCellData(String path,String sheetName,int rowNum,int colNum,String data,int firstRow,int lastRow,int firstCol,int lastCol) throws IOException {
		File xlFile=new File(path);
		if(!xlFile.exists()) {
			workbook=new XSSFWorkbook();
			fos=new FileOutputStream(path);
			workbook.write(fos);
		}
		if(workbook.getSheetIndex(sheetName)==-1) {
			workbook.createSheet(sheetName);
		}
		    sheet=workbook.getSheet(sheetName);
		if(sheet.getRow(rowNum)==null) {
			sheet.createRow(rowNum);
		}
		    row=sheet.getRow(rowNum);
		    cell=row.createCell(colNum);
		    cell.setCellValue(data);
		    sheet.addMergedRegion(new CellRangeAddress(firstRow,lastRow,firstCol,lastCol));
		    fos=new FileOutputStream(path);
		    workbook.write(fos);
		    workbook.close();
		    fos.close();
		    fis.close();	
	}
	
	
	
}
