Feature: Excel Read and Write

  Scenario: Write first name in Excel Sheet row 1 to 10
    Given I have an Excel file "C:\Users\soumyajit.pan\eclipse-workspace\ExcelReaderWriter\src\test\resources\Data\Employee.xlsx"
    Then I read and write data from sheet "Sheet1" for rows 1 to 10 for "FirstName"

	Scenario: Write first name in Excel Sheet row 11 to 20
    Given I have an Excel file "C:\Users\soumyajit.pan\eclipse-workspace\ExcelReaderWriter\src\test\resources\Data\Employee.xlsx"
    Then I read and write data from sheet "Sheet1" for rows 11 to 20 for "FirstName"

  Scenario: Write last name in Excel Sheet row 1 to 10
    Given I have an Excel file "C:\Users\soumyajit.pan\eclipse-workspace\ExcelReaderWriter\src\test\resources\Data\Employee.xlsx"
    Then I read and write data from sheet "Sheet1" for rows 1 to 10 for "LastName"

	Scenario: Write last name in Excel Sheet row 11 to 20
    Given I have an Excel file "C:\Users\soumyajit.pan\eclipse-workspace\ExcelReaderWriter\src\test\resources\Data\Employee.xlsx"
    Then I read and write data from sheet "Sheet1" for rows 11 to 20 for "LastName"
