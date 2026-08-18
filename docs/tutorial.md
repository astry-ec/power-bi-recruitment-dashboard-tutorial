# Build Your First Recruitment Dashboard in Power BI

## What you will build
In this tutorial, we will learn step by step how to build Power BI Dashboard with Excel file as the data source. Along the way we will also learn how to create DAX measures and calculated columns to get the required KPI data.
In addition, we will learn how to build some charts using several types of Power BI chart and creating slicer to narrow down the data.

## Prerequisites
Before we begin this tutorial, make sure you have:

- Power BI Desktop installed on your computer.
- The Excel file that will be used as the data source throughout this tutorial.
- Basic familiarity with Power BI, such as navigating the Power BI Desktop interface and using its main features.

### 1. Download the sample dataset
   
- Go to the data folder
- Locate the Candidate Data.xlsx file
- Download the file
<img width="1110" height="206" alt="image" src="https://github.com/user-attachments/assets/9d73232a-cc9c-47d8-8911-07da8910eb72" />


### 2. Import the data into Power BI

a. From "Home", find the "Get Data" and click Excel Workbook
<p>&nbsp;</p>
<img width="445" height="570" alt="image" src="https://github.com/user-attachments/assets/27e187e7-085e-4fd6-8a81-ca896ce71632" />
<p>&nbsp;</p>
b. Find the Candidate Data.xlsx in your local folder
<p>&nbsp;</p>
<img width="616" height="419" alt="image" src="https://github.com/user-attachments/assets/df6a23e3-6464-46bd-9231-aa2a23b28684" />
<p>&nbsp;</p>
c. The file will be loaded in Navigator pane as below. We can click the table name and check the data inside the table. Tick table Application, Candidate, Interview, Job, Source, StatusHistory.
  We don't need to change any data from the source, we just need to click "Load".
<p>&nbsp;</p>
<img width="888" height="704" alt="image" src="https://github.com/user-attachments/assets/2a6e9026-3935-4b6b-be56-c97a96a97318" />
<p>&nbsp;</p>
d. You will notice the data is loaded in the Data panel, in the right side.
<p>&nbsp;</p>
<img width="1920" height="519" alt="image" src="https://github.com/user-attachments/assets/8db7ce4f-99bf-4447-9a19-d7352971d343" />
<p>&nbsp;</p>

### 3. Prepare Additional Table

The best practice of data analytics relies on time-based visualization to show the trend over time. To achieve this, we need to create additional Calendar table.

a. From "Modeling" Pane, click the "New Measure" icon.
<p>&nbsp;</p>
<img width="946" height="341" alt="image" src="https://github.com/user-attachments/assets/2d5a9604-8021-4a8d-ab07-a12f73a7cb0e" />
<p>&nbsp;</p>
b. You can copy paste DAX formula below to create Calendar table.
<p>&nbsp;</p>
```
Calendar = 
VAR AllDates =
    UNION (
        SELECTCOLUMNS (
            'Application',
            "DateValue", 'Application'[ApplicationDate]
        ),
        SELECTCOLUMNS (
            'Job',
            "DateValue", 'Job'[OpenDate]
        ),
        SELECTCOLUMNS (
            'Job',
            "DateValue", 'Job'[CloseDate]
        ),
        SELECTCOLUMNS (
            'StatusHistory',
            "DateValue", 'StatusHistory'[StatusDate]
        )
    )

VAR MinDate =
    MINX (
        FILTER ( AllDates, NOT ISBLANK ( [DateValue] ) ),
        [DateValue]
    )

VAR MaxDate =
    MAXX (
        FILTER ( AllDates, NOT ISBLANK ( [DateValue] ) ),
        [DateValue]
    )

RETURN
    ADDCOLUMNS (
        CALENDAR ( MinDate, MaxDate ),
        "Year", YEAR ( [Date] ),
        "Month Number", MONTH ( [Date] ),
        "Month", FORMAT ( [Date], "MMMM" ),
        "Month Short", FORMAT ( [Date], "MMM" ),
        "Year Month", FORMAT ( [Date], "yyyy-MM" ),
        "Year Month Sort", YEAR ( [Date] ) * 100 + MONTH ( [Date] ),
        "Quarter", "Q" & FORMAT ( [Date], "Q" ),
        "Week Number", WEEKNUM ( [Date], 2 ),
        "Day", DAY ( [Date] ),
        "Day Name", FORMAT ( [Date], "dddd" )
    )

```
<p>&nbsp;</p>


5. 
6. Create your first DAX measures
7. Build KPI cards
8. Visualize applicant trends
9. Create a recruitment funnel
10. Add filters
11. Verify your dashboard

## Next steps
