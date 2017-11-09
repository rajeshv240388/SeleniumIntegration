package com.util.pack.main;

import java.io.IOException;

import com.util.pack.ExcelApiTest;

public class TestCase {
	
public static void main(String[] args) throws IOException {
	try {
		ExcelApiTest excel=new ExcelApiTest("D://Book1.xlsx");
		String value3=excel.getCellData("Credentials", "UserName", 2);
		String value4=excel.getCellData("Credentials", "PassWord", 2);
		String value5=excel.getCellData("Credentials", "DateCreated", 2);
		String value6=excel.getCellData("Credentials", "NoOfAttempts", 2);
		System.out.println(value3);
		System.out.println(value4);
		System.out.println(value5);
		System.out.println(value6);
		
	} catch (Exception e) {
		// TODO: handle exception
	}
	
}
}
