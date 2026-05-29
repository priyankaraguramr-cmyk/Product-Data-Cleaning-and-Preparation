# Product-Data-Cleaning-and-Preparation
Description
This assignment focuses on cleaning and preparing a retail product sales dataset using Microsoft Excel. The raw dataset had 34 records which was reduced to 32 rows after removing duplicate entries. The tasks covered handling missing values, correcting inconsistent text, removing duplicates, splitting and merging columns, applying currency and date formatting, and using conditional formatting to highlight patterns in the data.

Getting Started
Dependencies
•   Microsoft Excel 2016 or later
•   The raw dataset file: Excel_Assignment_2_Data_Exploration.xlsx

Files in this Repository
•   Excel_Assignment_2_Data_Exploration.xlsx — Raw input dataset
•   Excel_Assignment_2_Data_Exploration_Output.xlsx — Final cleaned and formatted dataset
•   Excel_Assignment2_Documentation.pdf — Step by step screenshots and explanations for all tasks

Tasks Completed
Task 1 – Handling Missing Values
•   Checked the Price column for missing values using Go To Special (Ctrl+G then Blanks)
•   Checked the Category column for blank cells using the same method
•   No missing values were found in either column
•   Strategy defined: infer Category from Product Name, use Uncategorised as fallback if not determinable

Task 2 – Correcting Inconsistent Data
•   Reviewed the Product Name column for text formatting inconsistencies
•   Found one issue: T-shirt in Row 7 used lowercase s against the Title Case standard
•   Corrected using Find and Replace (Ctrl+H): T-shirt replaced with T-Shirt
•   Reviewed the Category column for typos — all five values were correctly spelled, no changes needed

Task 3 – Removing Duplicate Records
•   Used Data then Remove Duplicates with all six columns selected
•   Found and removed two exact duplicate rows
•   Row 24 was a duplicate of Row 16 (21-AUG-CA, Laptop Bag, Samsonite, 0)
•   Row 33 was a duplicate of Row 11 (17-JUN-IN, Laptop, HP, 50)
•   Dataset reduced from 34 rows to 32 rows

Task 4 – Splitting and Merging Data
•   Split Product ID into Manufacturing Date (Column G) using the TEXT DATEVALUE formula
•   Split Product ID into Country Code (Column H) using the RIGHT formula
•   Created a new Product Brand column (Column I) by merging Brand Name and Product Name

Task 5 – Number Formatting
•   Formatted the Price column as currency using Format Cells (Ctrl+1 then Currency)
•   Applied format $#,##0.00 so values display like $1,000.00
•   Manufacturing Date column displays dates in DD-MM-YYYY format using the TEXT formula

Task 6 – Conditional Formatting
•   Applied Blue Data Bars to the Price column (D2:D33) for visual price comparison
•   Created a custom rule on the Category column to highlight Electronics rows in yellow with bold text

Formulas Used

Manufacturing Date   =TEXT(DATEVALUE(LEFT(A2,6)&"-2024"),"DD-MM-YYYY")
Country Code         =RIGHT(A2,2)
Product Brand        =C2&" "&B2

Dataset Columns

Column Name	Description
Product ID	Unique identifier combining manufacturing date and country code (e.g. 28-JAN-US)
Product Name	Name of the product like Laptop, Sneakers, Blender etc.
Brand Name	Brand of the product like Dell, Nike, Ninja etc.
Price ($)	Price of the product in US dollars
Quantity	Number of units sold
Category	Product category — Electronics, Fashion, Kitchen, Outdoor or Accessories
Manufacturing Date	Date extracted from Product ID in DD-MM-YYYY format — added in Column G
Country Code	Country code extracted from Product ID — added in Column H
Product Brand	Brand Name and Product Name merged into one column — added in Column I

Key Results
•   2 duplicate rows removed — dataset reduced from 34 to 32 rows
•   1 text inconsistency corrected — T-shirt changed to T-Shirt
•   No missing values found in any column
•   3 new columns added — Manufacturing Date, Country Code, Product Brand
•   Currency formatting applied to Price column
•   Conditional formatting applied to Price and Category columns



