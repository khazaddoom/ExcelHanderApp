package controller;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelSheetProcessor {
	
	
	public static Workbook createWorkbook(FileInputStream fis, String fileName) throws IOException {
		
		
		if (fileName.endsWith(".xlsx") || fileName.endsWith(".xlsm")) {
			
			return new XSSFWorkbook(fis);
			
		} else if (fileName.endsWith(".xls")) {

			return new HSSFWorkbook(fis);
			
		} else {
			
			throw new IllegalArgumentException("Incoming file name is not a supported type");

		}
		
		
	}
	
	
	

}
