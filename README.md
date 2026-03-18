# excel-hr-vacation-tracker
Excel based HR Leave Management &amp; Vacation Tracking System

Download file here:   [HR_vacation_tracker.xlsm](https://github.com/user-attachments/files/26086443/HR_vacation_tracker.xlsm)

---
**Key features**

    •	Automated working day vacation, sick day or personal day calculation

    •	Holiday and Weekend exclusion using NETWORKDAYS

    •	Return to workday using WORKDAY

    •	Vacation overlap detection

    •	Conditional formatting alerts

    •	Sheet protection to prevent formula edits

    •	Optional macro for safe row insertion in protected table

    •	Overview of all absent days using FILTER, XLOOKUP, SUMIFS
  
---
**System structure**

 Vacation input Table
 
       - recording individual vacation/sick day/personal day 
       - support multiple entries by employee
       - enter only start day and end day 
       - automatically calculates absent days and back to work day (excluding weekends and holidays)

 <img width="629" height="292" alt="Vacation tracker" src="https://github.com/user-attachments/assets/b2f0549d-fbc4-4fc9-aea5-1691b3f51eb5" />

---

Overview 

         - absent days by employee    
         - used vacation days alert

<img width="482" height="198" alt="Overview of absent days" src="https://github.com/user-attachments/assets/ecd75ff7-3866-4ee4-a263-e922b6a6265d" />

---

Employee Vacation days per Year

    - yearly vacation allowance per employee
    
<img width="247" height="337" alt="Employee Vacation days per Year" src="https://github.com/user-attachments/assets/936b34c4-8584-4fb8-a093-4869683be5c9" />

---

Dashboard

    - total vacation days by department  
    - used vacation days by employee  
    - total absent days by employee  
    - total vacation days by month  
    - department, year, month slicer

    
<img width="689" height="348" alt="Dashboard" src="https://github.com/user-attachments/assets/ce43d657-005d-4512-ad5d-81945852b898" />

---
    
Employees database & Databases

    - employee list  
    - holidays list  
    - other informations for data validation


<img width="211" height="113" alt="Employee Database" src="https://github.com/user-attachments/assets/f9363c08-49b2-4bab-b8a5-47279d34e05b" />

 <img width="407" height="295" alt="Other Databases" src="https://github.com/user-attachments/assets/b25bbbd6-235a-4878-856d-571077e5a626" />

---   

**Excel functions used**

    NETWORKDAYS
  
    WORKDAY
  
    XLOOKUP
  
    FILTER
  
    SUMIFS
  
    Stuctured Tables
  
    Conditional Formatting
  
    Data Validation
---
**VBA Usage**
    
    This project include s small VBA procedure 
    to improve using all table futures with protected Excel worksheets.

    When worksheet is protected using TAB to add new row may not function correctly. 
    This can prevent automatic formula propagation.

    The macro resolves this by:
        - Temporarily unprotecting the sheet
        - Adding a new row to the table
        - Reapplying protection

    All business logic and calculations remain implemented using native Excel formulas to ensure transparency  
  
  <img width="629" height="292" alt="Vacation tracker" src="https://github.com/user-attachments/assets/2ec55206-27c5-48b0-b611-31cd9b82d4d1" />

  ---

**File**

Download the Excel file here:   [HR_vacation_tracker.xlsm](https://github.com/user-attachments/files/26086443/HR_vacation_tracker.xlsm)

