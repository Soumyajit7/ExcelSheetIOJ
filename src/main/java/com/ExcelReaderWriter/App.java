package com.ExcelReaderWriter;


public class App {
	public static String filePath = "src\\test\\resources\\Data\\NewTest.xlsx";

	public static void main(String[] args) {
		String currentDirectory = System.getProperty("user.dir");
		String absoluteFilePath = currentDirectory + "\\" + filePath;
		System.out.println("Absolute File Path: " + absoluteFilePath);
		
		ExcelReaderWriter e = new ExcelReaderWriter(filePath, "EEID");
		String fullName = e.getTestData("Sheet1", "E3", "Full Name");
		System.out.println("Full Name -> " + fullName);
		
		// e.setTestData("Sheet1", "E3", "Full Name", "Soumyajit Pan");
		// fullName = e.getTestData("Sheet1", "E3", "Full Name");
		// System.out.println("Full Name -> " + fullName);
	}
}
