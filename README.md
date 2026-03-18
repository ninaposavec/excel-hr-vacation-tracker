# excel-hr-vacation-tracker
Excel based HR Leave Management &amp; Vacation Tracking System

**Key features**

    Automated working day vacation, sick day or personal day calculation

    Holiday and Weekend exclusion using NETWORKDAYS

    Return to workday using WORKDAY

    Vacation overlap detection

    Conditional formatting alerts

    Sheet protection to prevent formula edits

    Optional macro for safe row insertion in protected table

    Overview of all absent days using FILTER, XLOOKUP, SUMIFS
  

**System structure**

Vacation tracker Table
 
    individual vacation/sick day/personal day request

Overview 

    absent days by employee
    
    used vacation days alert

Employee Vacation days per Year

    yearly vacation allowance per employee

Dashboard

    total vacation days by department
  
    used vacation days by employee
  
    total absent days by employee
  
    total vacation days by month
  
    department, year, month slicer

Employees database & Databases

    employee list
  
    holidays list
  
    other informations for data validation
    

**Excel functions used**

    NETWORKDAYS
  
    WORKDAY
  
    XLOOKUP
  
    FILTER
  
    SUMIFS
  
    Stuctured Tables
  
    Conditional Formatting
  
    Data Validation

**VBA Usage**
    
    This project include s small VBA procedure to improve using all table futures with protected Excel worksheets.

    When worksheet is protected using TAB to add new row may not function correctly. This can prevent automatic formula          propagation.

    The macro resolves this by:
        - Temporarily unprotecting the sheet
        - Adding a new row to the table
        - Reapplying protection

All business logic and calculations remain implemented using native Excel formulas to ensure transparency and maintainability. 

**Screenshots**
  
  <img width="629" height="292" alt="Vacation tracker" src="https://github.com/user-attachments/assets/2ec55206-27c5-48b0-b611-31cd9b82d4d1" />

  
  




 
<img width="689" height="348" alt="Dashboard" src="https://github.com/user-attachments/assets/cda0caee-7d29-4d5b-bb83-a7802ac82e69" />



**File**

Download the Excel file here:   [HR_vacation_tracker.xlsm](https://github.com/user-attachments/files/26086443/HR_vacation_tracker.xlsm)

