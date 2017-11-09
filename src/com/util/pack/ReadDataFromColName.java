package com.util.pack;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;

public class ReadDataFromColName {
	static XSSFRow row;
public static void main(String[] args) throws EncryptedDocumentException, InvalidFormatException, IOException {
	
	FileInputStream stream=new FileInputStream(new File("D://Book1.xlsx"));
	Workbook book=WorkbookFactory.create(stream);
	Sheet sheet=book.getSheet("Credentials");
	row=(XSSFRow) sheet.getRow(0);
	int col_num=-1;
	for (int i = 0; i < row.getLastCellNum(); i++) {
		if (row.getCell(i).getStringCellValue().trim().equalsIgnoreCase("password")) {
			col_num=i;
		}
	}
	for (int j = 1; j < sheet.getLastRowNum()+1; j++) {
		row=(XSSFRow) sheet.getRow(j);
		XSSFCell cell=row.getCell(col_num);
	
	
	System.out.println((cell.getStringCellValue()) );
		
	}
	
	
		
	
}
}
