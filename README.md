# **Context**

You are working for a global digital agency that delivers digital solutions and products across multiple industries. Currently, the company manages numerous projects and tasks using scattered Excel spreadsheets. To improve project management, sales, marketing, and overall team collaboration, they are migrating to Salesforce, a leading customer success platform.

# **Situation**

The company has exported all past, current, and future project data to facilitate the migration. However, management lacks a clear overview of:

- Accounts and their statuses
- Active Projects and their related Tasks
- Completed Tasks
- Project Managers assigned to specific Projects

To ensure that only relevant and accurate data is migrated, the management team has requested a **business analysis**. As a **Data Business Analyst**, your role is to analyze the Excel data and provide actionable recommendations to key business stakeholders.

# **Business Analysis**

The provided spreadsheet contains data on **hundreds of companies and projects**. It consists of three key tabs:

## **1. Account Data**

- Contains details of all companies that the business works with, including those with **active, inactive, or canceled** projects.
- **Column A**: Company Name and Location
- **Column B**: Industry Sector
- **Column C**: Project Status (Active, Inactive, or Canceled)
- **Data Issue**: Some Accounts are **duplicated** due to poor data quality and missing duplicate checks.

## **2. Project & Task Data**

- Contains details of **active projects** and their associated tasks.
- **Column A**: Unique Account Name
- **Column B**: Active Project Status (Yes/Pending for active, No/Canceled for inactive)
- **Column C**: Project Start Year
- **Column D**: Number of Tasks per Project
- **Column E**: Overall Task Status

## **3. Project Manager Data**

- Contains details of **active projects** and their assigned Project Managers.
- **Project Manager assignments** are based on the **Industry sector** of the Account (as per an assignment table).

Using this data, please answer the following questions. We will evaluate both the **accuracy of your answers** and the **efficiency of your data analysis methods**.

---

# **Assignment**

## **Analysis**

1. How many unique Accounts exist?
2. Which Industry sectors are represented by these Accounts? What is the largest Industry sector?
3. How many Accounts and Projects are **active**? _(Active projects are marked as "Yes" or "Pending")_
4. How many Accounts and Projects are **inactive**? _(Marked as "No" or "Canceled")_
5. How many **Active Projects** exist per year?
6. How many **Tasks** are there per Project?
7. How many Projects are assigned to each **Project Manager**?

## **Rationale**

8. What **assumptions** did you make in your analysis?
9. Is there any **missing data** that should be included in the spreadsheet?
10. What **recommendations** would you give to ensure **data quality and integrity** during the Salesforce migration?

---

# **Required Files**

Since the Excel file contains **confidential information**, it is not included in this repository. This repository serves as a reference for the **Excel macros** I created for similar future use cases.
