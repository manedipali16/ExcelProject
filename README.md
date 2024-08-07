# ExcelProject
Assignment 1: Macro for Analysis
Scenario:
You are given an Excel file containing insurance premiums for multiple customers.

Steps:
Open the Excel file (Insurance.xlsx).

Record a Macro:

Go to the View tab and click on Macros, then choose Record Macro.
Name your macro (e.g., InsuranceAnalysisMacro) and choose where to store it (e.g., This Workbook).
Create a New Column for Total Insured:

Select the first empty column and name it Total Insured.
Enter the number of people covered under each policy manually or through a formula if provided.
Calculate Premium Paid Per Head:

Create another new column next to Total Insured and name it Premium Per Head.
Enter the formula to calculate the premium per head:
excel
Copy code
= [Expense Column] / [Total Insured Column]
Drag the formula down to apply it to all rows.
Stop the Macro:

Go to the View tab, click on Macros, and choose Stop Recording.
Save the File:

Save the Excel file with the macro enabled (.xlsm format).
Assignment 2: Data Validation
Scenario:
You are given an Excel file containing insurance premiums for multiple customers.

Steps:
Open the Excel file (Insurance.xlsx).

Apply Data Validation:

Age Column:

Select the cells in the Age column.
Go to the Data tab, click Data Validation.
Set Allow to Whole number, Data to between, Minimum to 18, and Maximum to 80.
Gender Column:

Select the cells in the Gender column.
Go to the Data tab, click Data Validation.
Set Allow to List, and Source to male, female, others.
BMI and Expenses Columns:

Select the cells in the BMI column, repeat for the Expenses column.
Go to the Data tab, click Data Validation.
Set Allow to Decimal or Whole number, and Data to greater than, with Minimum set to 0.
Smoker Column:

Select the cells in the Smoker column.
Go to the Data tab, click Data Validation.
Set Allow to List, and Source to yes, no.
Save the File:

Save your changes to the Excel file.
Assignment 3: Dashboard Creation
Scenario:
You are given an Excel file containing insurance premiums for multiple customers.

Steps:
Open the Excel file (Insurance.xlsx).

Prepare the Data:

Make sure your data is clean and organized, with relevant columns like Age, Gender, Smoker, Region, Expense, etc.
Create a Line Chart:

Objective: Depict the relation between the average premium per head and age.
Calculate the average premium per head for each age group.
Select the data, go to the Insert tab, and choose a Line Chart.
Create a Bar Chart:

Objective: Show the fraction of the male and female population who smoke.
Create a table showing the count of smokers by gender.
Select the data, go to the Insert tab, and choose a Bar Chart.
Create a Pie Chart:

Objective: Depict the count of people who smoke in each region.
Create a table summarizing the count of smokers in each region.
Select the data, go to the Insert tab, and choose a Pie Chart.
Arrange the Charts in a Dashboard:

Place the charts on a single sheet, arranging them to create a cohesive dashboard.
Use Slicers or Pivot Tables if necessary to add interactivity.
