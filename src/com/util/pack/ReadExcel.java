package com.util.pack;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ReadExcel {
static Sheet sheet;
	public static String[][] getcelldata(String Path, String sheetname) throws EncryptedDocumentException, InvalidFormatException, IOException{
		FileInputStream stream=new FileInputStream(new File(Path));
		Workbook book=WorkbookFactory.create(stream);
		Sheet sheet=book.getSheet(sheetname);
		int rowcount=sheet.getLastRowNum();
		int cellcount=sheet.getRow(0).getLastCellNum();
		String data[][]=new String[rowcount][cellcount];
		for (int i = 0; i <rowcount; i++) {
			Row row=sheet.getRow(i);
			for (int j = 0; j < cellcount; j++) {
				Cell c=row.getCell(j);
				try {
					if (c.getCellTypeEnum()==CellType.STRING) {
						data[i][j]=c.getStringCellValue();
					} else if (c.getCellTypeEnum()==CellType.BLANK) 
                    return null;
					 else 
						data[i][j]=String.valueOf(c.getNumericCellValue());
					
				} catch (Exception e) {
					
				}
			}
		}
	return data;	
	}
	public static void main(String[] args) throws EncryptedDocumentException, InvalidFormatException, IOException {
		try {
			String value[][]=getcelldata("D://Book1.xlsx", "max");
			System.out.println(value.length);
			System.out.println(value[0].length);
			for (int k = 0; k <value.length ; k++) {
				for (int l = 0; l <value[0].length; l++) {
					System.out.println(value[k][l]);
				}
			}
		} catch (NullPointerException nullPointer) {
			// TODO: handle exception
		}catch (Exception e) {
			e.printStackTrace();
		}
		

	}

}
