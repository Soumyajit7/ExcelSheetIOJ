/**************************************
* Developed By : Soumyajit Pan
* Updated Date : 22 January, 2024
**************************************/

package com.ExcelReaderWriter;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.*;
import java.io.*;
import java.nio.file.*;

public class ExcelReaderWriter {
	private String originalFilePath;
	private Workbook wb;
	private Map<String, Map<String, Map<String, Cell>>> data;

	public ExcelReaderWriter(String excelFilePathWithName, String primaryKeyColName) {
		this.originalFilePath = excelFilePathWithName;
		this.data = new HashMap<>();
		try {
			FileInputStream fis = new FileInputStream(new File(this.originalFilePath));
			this.wb = new XSSFWorkbook(fis);
			for (Sheet sheet : wb) {
				Map<String, Map<String, Cell>> sheetData = new HashMap<>();
				Row headerRow = sheet.getRow(0);
				int primaryKeyColIndex = -1;
				// Find the index of the primary key column
				for (Cell cell : headerRow) {
					if (cell.getStringCellValue().equals(primaryKeyColName)) {
						primaryKeyColIndex = cell.getColumnIndex();
						break;
					}
				}
				if (primaryKeyColIndex == -1) {
					throw new IllegalArgumentException("Primary key column '" + primaryKeyColName + "' not found");
				}
				for (Row row : sheet) {
					if (row.getRowNum() == 0)
						continue; // Skip header row
					Map<String, Cell> rowData = new HashMap<>();
					for (int j = 0; j < headerRow.getLastCellNum(); j++) {
						String colName = headerRow.getCell(j).getStringCellValue();
						Cell cell = row.getCell(j, Row.MissingCellPolicy.RETURN_NULL_AND_BLANK);
						if (cell == null || cell.getCellType() == CellType.BLANK) {
							cell = row.createCell(j);
							cell.setCellValue("NULL");
						}
						rowData.put(colName, cell);
					}
					String primaryKeyValue = row.getCell(primaryKeyColIndex).getStringCellValue();
					sheetData.put(primaryKeyValue, rowData);
				}
				this.data.put(sheet.getSheetName(), sheetData);
			}
			fis.close();
			// printData(this.data);
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	@SuppressWarnings("unused")
	private void printData(Map<String, Map<String, Map<String, Cell>>> data2) {
		for (Map.Entry<String, Map<String, Map<String, Cell>>> entry : data2.entrySet()) {
			String key1 = entry.getKey();
			Map<String, Map<String, Cell>> value1 = entry.getValue();

			for (Map.Entry<String, Map<String, Cell>> entry2 : value1.entrySet()) {
				String key2 = entry2.getKey();
				Map<String, Cell> value2 = entry2.getValue();

				for (Map.Entry<String, Cell> entry3 : value2.entrySet()) {
					String key3 = entry3.getKey();
					Cell cell = entry3.getValue();

					System.out.println("Key1: " + key1 + ", Key2: " + key2 + ", Key3: " + key3 + ", Cell: " + cell);
				}
			}
		}
	}

	public String getTestData(String sheetName, String testCaseName, String colName) {
		/**
		 * This method retrieves test data from the stored Excel data.
		 * Here's how it works:
		 * 1. It takes the sheet name, test case name, and column name as input.
		 * 2. It uses these parameters to access the corresponding cell in the data
		 * HashMap.
		 * 3. It returns the string value of the cell.
		 * This method allows for easy retrieval of specific pieces of test data from
		 * the Excel file.
		 * 
		 * @param sheetName    The name of the sheet in the Excel file.
		 * @param testCaseName The name of the test case (usually corresponds to a row
		 *                     in the sheet).
		 * @param colName      The name of the column in the sheet.
		 * @return The string value of the specified cell.
		 */
		Cell cell = this.data.get(sheetName).get(testCaseName).get(colName);
		return cell.getStringCellValue();
	}

	public void setTestData(String sheetName, String testCaseName, String colName, String writableData) {
		/**
		 * This method writes data to a specified cell in the Excel file.
		 * Here's how it works:
		 * 1. It creates a temporary copy of the original Excel file.
		 * 2. It opens the temporary file and locates the specified sheet and row.
		 * 3. It finds the column index based on the column name.
		 * 4. It ensures the cell exists before setting its value to the provided data.
		 * 5. It saves the changes to the temporary file.
		 * 6. It copies the temporary file back to the original file, effectively saving
		 * the changes to the original file.
		 * 7. It deletes the temporary file.
		 * 8. Finally, it updates the data attribute to reflect the changes made to the
		 * Excel file.
		 * This method allows for easy modification of specific pieces of test data in
		 * the Excel file.
		 * 
		 * @param sheetName    The name of the sheet in the Excel file.
		 * @param testCaseName The name of the test case (usually corresponds to a row
		 *                     in the sheet).
		 * @param colName      The name of the column in the sheet.
		 * @param writableData The data to write to the specified cell.
		 */
		try {
			// Create a temporary file
			File tempFile = new File(getCopyFolderDirectory(), getTempFileName(getFileName(this.originalFilePath)));
			System.out.println("tempFile->" + tempFile);
			String tempFilePath = tempFile.getAbsolutePath();

			// Copy original file to temporary file
			Files.copy(Paths.get(this.originalFilePath), Paths.get(tempFilePath), StandardCopyOption.REPLACE_EXISTING);

			// Load workbook from temporary file
			FileInputStream fis = new FileInputStream(tempFilePath);
			Workbook wb = new XSSFWorkbook(fis);
			fis.close();

			// Get sheet and max row
			Sheet sheet = wb.getSheet(sheetName);
			int maxRow = sheet.getLastRowNum();

			// Iterate over rows
			for (int i = 1; i <= maxRow; i++) {
				Row row = sheet.getRow(i);
				Cell cell = row.getCell(0);
				String testName = cell.getStringCellValue();

				// Check if test case name matches
				if (testCaseName.equals(testName)) {
					int colIndex = 0;
					Row headerRow = sheet.getRow(0);

					// Find column index
					while (!colName.equals(headerRow.getCell(colIndex).getStringCellValue())) {
						colIndex++;
					}

					// Ensure cell exists before setting its value
					Cell targetCell = row.getCell(colIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
					targetCell.setCellValue(writableData);
					break;
				}
			}

			// Save workbook to temporary file
			FileOutputStream fos = new FileOutputStream(tempFilePath);
			wb.write(fos);
			fos.close();
			wb.close();

			// Copy temporary file back to original file
			Files.copy(Paths.get(tempFilePath), Paths.get(this.originalFilePath), StandardCopyOption.REPLACE_EXISTING);

			// Delete temporary file
			tempFile.delete();

			// Update data attribute
			updateData(this.data, sheetName, testCaseName, colName, writableData);

		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	private static String getFileName(String filePath) {
		/**
		 * This method retrieves the filename from a given file path.
		 * Here's how it works:
		 * 1. It converts the string file path into a Path object.
		 * 2. It then uses the getFileName method of the Path class to retrieve the
		 * filename.
		 * This method is useful when you have a file's full path and want to isolate
		 * just the filename.
		 * 
		 * @param filePath The full path of the file.
		 * @return The filename as a string.
		 */
		Path path = Paths.get(filePath);
		return path.getFileName().toString();
	}

	private String getTempFileName(String filename) {
		/**
		 * This method generates a temporary filename by appending a random number to
		 * the base name of the original filename.
		 *
		 * @param filename The original filename.
		 * @return The temporary filename.
		 */
		String baseName = filename.substring(0, filename.lastIndexOf("."));
		String extension = filename.substring(filename.lastIndexOf("."));
		int randomNumber = new Random().nextInt(1000) + 1;
		return baseName + "_temp" + randomNumber + extension;
	}

	public void removeTemporaryDataFiles() {
		/**
		 * This method deletes all temporary data files in the "CopyFolder" directory.
		 * Here's how it works:
		 * 1. It gets the directory of the current class file.
		 * 2. It constructs the path to the "CopyFolder" directory within the class
		 * file's directory.
		 * 3. It lists all files in the "CopyFolder" directory.
		 * 4. It iterates over each file in the directory.
		 * 5. If a file's name starts with "temp", it attempts to delete the file.
		 * 6. If the file is successfully deleted, it prints a confirmation message. If
		 * the file cannot be deleted, it prints an error message.
		 * This method is useful for cleaning up temporary files that are no longer
		 * needed.
		 */
		String fileDir = new File(ExcelReaderWriter.class.getProtectionDomain().getCodeSource().getLocation().getPath())
				.getParent();
		String myDir = Paths.get(fileDir, "CopyFolder").toString();
		File folder = new File(myDir);
		File[] listOfFiles = folder.listFiles();

		for (File file : listOfFiles) {
			if (file.isFile() && file.getName().startsWith("temp")) {
				try {
					if (file.delete()) {
						System.out.println("Deleted the file: " + file.getName());
					} else {
						System.out.println("Failed to delete the file.");
					}
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		}
	}

	@SuppressWarnings("unused")
	private void saveAndClose() {
		/**
		 * This method saves and closes the workbook.
		 * Here's how it works:
		 * 1. It generates a temporary filename using the getTempFileName method.
		 * 2. It creates a FileOutputStream for the temporary file and writes the
		 * workbook to this stream.
		 * 3. It closes the FileOutputStream.
		 * 4. It copies the temporary file back to the original file, effectively saving
		 * the changes to the original file.
		 * 5. It deletes the temporary file.
		 * This method is useful for saving changes to the workbook without risking data
		 * loss or corruption in the original file.
		 * 
		 * @throws IOException If an I/O error occurs.
		 */
		try {
			// Save the workbook to a temporary file
			String tempFileName = getTempFileName(this.originalFilePath);
			FileOutputStream fos = new FileOutputStream(new File(tempFileName));
			this.wb.write(fos);
			fos.close();

			// Copy the temporary file back to the original file
			Files.copy(Paths.get(tempFileName), Paths.get(this.originalFilePath));

			// Delete the temporary file
			Files.delete(Paths.get(tempFileName));
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	private String getCopyFolderDirectory() {
		/**
		 * This method gets the directory of a folder named "CopyFolder" in the current
		 * working directory.
		 * If the folder does not exist, it creates one.
		 *
		 * @return The directory of the "CopyFolder".
		 */
		String currentDirectory = System.getProperty("user.dir");
		Path copyFolderPath = Paths.get(currentDirectory, "CopyFolder");
		if (!Files.exists(copyFolderPath)) {
			try {
				Files.createDirectory(copyFolderPath);
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
		return copyFolderPath.toString();
	}

	public List<String> readAllDataInGivenColumn(String sheetName, String colName) {
		/**
		 * This method reads all data in a given column from a specified sheet.
		 *
		 * @param sheetName The name of the sheet.
		 * @param colName   The name of the column.
		 * @return A list of data in the specified column. If the sheet or column does
		 *         not exist, it returns a list with an appropriate message.
		 */
		if (!this.data.containsKey(sheetName)) {
			return Collections.singletonList("Sheet not found");
		}

		List<String> columnData = new ArrayList<>();
		for (Map<String, Cell> testCase : this.data.get(sheetName).values()) {
			if (testCase.containsKey(colName) && testCase.get(colName).getStringCellValue() != null) {
				columnData.add(testCase.get(colName).getStringCellValue());
			}
		}

		if (columnData.isEmpty()) {
			return Collections.singletonList("No data found in column " + colName + " in sheet " + sheetName);
		}

		return columnData;
	}

	private void updateData(Map<String, Map<String, Map<String, Cell>>> data, String sheetName, String testCaseName,
			String colName, String newCellValue) {
		/**
		 * This method updates the data in the specified cell.
		 * Here's how it works:
		 * 1. It retrieves the sheet, test case, and cell from the data HashMap using
		 * the provided names.
		 * 2. If the sheet, test case, or cell does not exist, it prints an error
		 * message.
		 * 3. If all exist, it sets the cell's value to the new value.
		 * This method allows for easy modification of specific pieces of data in the
		 * HashMap.
		 * 
		 * @param data         The HashMap storing the Excel data.
		 * @param sheetName    The name of the sheet in the Excel file.
		 * @param testCaseName The name of the test case (usually corresponds to a row
		 *                     in the sheet).
		 * @param colName      The name of the column in the sheet.
		 * @param newCellValue The new value to set in the specified cell.
		 */
		Map<String, Map<String, Cell>> sheet = data.get(sheetName);
		if (sheet != null) {
			Map<String, Cell> testCase = sheet.get(testCaseName);
			if (testCase != null) {
				Cell cell = testCase.get(colName);
				if (cell != null) {
					cell.setCellValue(newCellValue);
				} else {
					System.out.println("Column Name: " + colName + " does not exist.");
				}
			} else {
				System.out.println("Test Case Name: " + testCaseName + " does not exist.");
			}
		} else {
			System.out.println("Sheet Name: " + sheetName + " does not exist.");
		}
	}

}
