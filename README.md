# MS bot framework functional testing utility

This repository contains the basic code to start writing the functional regression test suite for chat bot written using MS Bot Framework.
This can be enhanced further to have a fancy UI and robust reporting.
#### Code files
- TestSuite.cs: Reads the provided excel file and create the functional test cases flow
- Helper: Helper files to read the excel file
  - ObjectExtensions.cs
  - UriFixer.cs
  - WorksheetExtension.cs
- Program.cs: Entry file for the project

#### Structure of excel file (Test case file)
Tabs in excel
- Index: contains information for all test cases in current test suite
Columns in tab
  - TestCaseNumber: Name of the test case workbook sheet
  - Description: Brief description about the test case
![Image of index tab](https://raw.githubusercontent.com/lokeshlal/MS-Bot-Test/master/MSBotTest/index_tab.PNG)
- Test case tabs (TC1, TC2): contains test case flow
Columns in tab
  - Step: test case step number 
  - User: user input
  - Response: expected bot response
  - Entity: Entities in user input, which can be used in further steps for validation of response
![Image of test case tab](https://raw.githubusercontent.com/lokeshlal/MS-Bot-Test/master/MSBotTest/testcase_tab.PNG)

Entities are index based, for example "{0,Lokesh}" means entity with index 0 and content as "Lokesh" or "{0,Lokesh}{1,Lucky}" means 2 entity, one at index 0 with content "Lokesh" and another at index 1 with content "Lucky"

To use entities in bot response validation, use ${stepNumber} to insert the complete user response at step-number. or use ${stepNumber-index} that us use entity from stepNumber at index.

For example, ${2} will be replaced with "Lokesh"/"Lucky" depending upon the flow. Similarly, ${2-0} will replace "Lok"/"Lucky" depending upon the flow.
