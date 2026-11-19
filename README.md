
# Project 1: Retail Sales Data Cleaning Using Microsoft Excel

## Purpose
Cleaning Data

## Tools Used
- Microsoft Excel

## Business Objective

The objective of this project is to transform poorly presented retail sales records into a well-structured, consistent, and professionally formatted spreadsheet using Microsoft Excel.

The task involves improving the readability and presentation of the dataset through appropriate cell formatting, column alignment, number and date formatting, borders, headings, and column adjustments.

The aim is to make the sales records easier for the company's management team to read, review, and use for future business reporting and analysis, while preserving the original data.

## Project Background: How I Started

Pretending that I was working for a retail company, I decided to take on my first Excel project focused on data cleaning.

How did I approach it?

I created a prompt using ChatGPT to generate rough sales data, putting myself in the position of a data analyst working for a retail company. I imagined that the manager wanted the job completed before the next meeting.

The rough dataset contained Order ID, Customer Name, Product, Region, Quantity, Unit_Price, Order_Price, and Payment_Status, with different types of data issues that I had to work on.

### Original Rough Data

![Original Rough Data](Screenshots/01-raw-data.png)

## Step 1: Understanding the Dataset

After getting access to the data, I duplicated the original dataset to have a copy to work with. I then took my time to look through the data carefully to identify what was wrong and how I could turn the rough data into something meaningful.

The first thing I noticed was that the table did not have a title. I checked my practice checklist and decided to name it **Monthly Sales Report – March 2026**.

After that, I manually adjusted the columns and their widths to make the contents readable and properly aligned. I also formatted the column headers by making them bold and giving them a different colour to make them stand out.

## Step 2: Cleaning Customer Names, Products and Regions

Moving on to Customer_Name, Product, and Region, I noticed that I needed to maintain consistency between uppercase and lowercase letters.

I used the TRIM and PROPER functions to clean up extra spaces and make the text presentation more consistent.

I paid particular attention to the Region column because Excel may treat the same region as different entries when they are written in different ways. Maintaining consistency was important for future reporting and analysis.

## Step 3: Formatting Unit Price and Order Date

Moving on to Unit_Price, I used Format Cells to add the Nigerian Naira (₦) currency symbol. I also changed the number format from General to Currency to give the prices a better presentation.

I did something similar with the Order_Date column by making sure the dates followed a consistent structure using Format Cells.

## Step 4: Cleaning Payment Status

Looking at the Payment_Status column, I noticed that there were different entries, including Paid, Pending, Not Paid, and Unknow.

I wanted to make the payment status easier to understand and maintain consistency throughout the dataset.

I used TRIM, UPPER, and PROPER to handle text formatting and remove unwanted spaces. I also used Conditional Formatting to apply different colours to the payment statuses, making them easier to identify.

## Step 5: Calculating Total Sales

I also decided to add a Total_Sales column to the table. I thought this would be useful for the team when reviewing the monthly sales report.

I used a simple multiplication formula:

`=E4*F4`

After entering the formula, I dragged it down to apply the calculation to the remaining records.

This helped calculate the total sales value for each order based on its quantity and unit price.

## Step 6: Identifying Data Quality Issues

While working on the dataset, I noticed that the Quantity column had some missing values, while some records contained negative numbers. I also discovered that one of the dates in the Order_Date column was missing.

I needed to find a way to document these issues and explain the decisions I made while working on them.

## Step 7: Creating a Data Issue Log

To handle this, I created a Data_Issue_Log to explain my findings and the actions I took to the team and the manager.

The log contained:

- Order ID
- Issue Found
- Action Taken
- Resolved

This allowed me to document the problems I discovered, explain how I handled them, and identify the issues that still required further confirmation.

### Data Issue Log

![Data Issue Log](Screenshots/04-data-issue-log.png)

## Step 8: Creating a Summary for the Team

I also created a summary sheet to give the team a quick overview of the data and make it easier to understand what was happening in the dataset.

The idea was to provide useful information that could help the team review the records and make better decisions in the future.

### Data Quality Summary

![Data Quality Summary](Screenshots/05-summary-sheet.png)

## Final Result

After working through the different stages, I was able to transform the original rough sales data into a more organised and professionally presented spreadsheet.

![Final Cleaned Dataset](Screenshots/03-cleaned-data.png)

## What I Have Learned So Far

Through this project, I have gained practical experience in:

- Creating professional titles for Excel reports.
- Using the Format Cells section in Excel.
- Converting data ranges into Excel Tables.
- Adjusting column widths and aligning data.
- Applying Conditional Formatting.
- Using TRIM, PROPER, and UPPER functions.
- Formatting currency and dates.
- Using formulas to perform calculations.
- Identifying missing and incorrect values.
- Creating a Data Issue Log to document data problems.
- Making decisions based on available evidence.
- Presenting data in a way that makes it more useful for a team.

## Conclusion

This project gave me an opportunity to understand that data cleaning is not just about making a spreadsheet look beautiful. It also involves paying attention to details, identifying problems, making careful decisions, and documenting the actions taken.

It was my first practical Excel data cleaning project, and it helped me understand how Microsoft Excel can be used to prepare data for business reporting and future analysis.

