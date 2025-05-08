```
Name : PRAKASH C
Reg.No : 212223240122
```
# RPA_Ex-04
# Aim:
  To create a UiPath workflow that reads data from one Excel file and writes it to another Excel file using Excel activities.

# Materials Required:
 1. Uipath Studio
 2. Microsoft Excel
 
# PROCEDURE:

## Step 1:
Create a New Process Open UiPath Studio and create a new process named ExcelReadWriteExample.
## Step 2: 
Prepare Excel Files Create an Excel file named Input.xlsx with sample data, e.g.:
![Screenshot 2025-05-08 103736](https://github.com/user-attachments/assets/357d9cfa-5448-4cfd-9503-df8dbc021418)

Save this file in your project directory.

## Step 3: 
Add Excel Application Scope Drag and drop Excel Application Scope activity.
Set the path to "Input.xlsx".

## Step 4: 
Read Range Activity Inside the Excel Scope, drag a Read Range activity.
Properties: SheetName: "Sheet1" Range: Leave empty to read the whole sheet Output: inputDataTable (Create a variable of type DataTable)

## Step 5:
Add Second Excel Application Scope Below the first scope, drag another Excel Application Scope.
Set the path to "Output.xlsx" (can be a new empty file).

## Step 6: 
Write Range Activity Inside the second scope, drag a Write Range activity.
Properties: DataTable: inputDataTable SheetName: "Sheet1" Starting Cell: "A1" Add Headers: ✔️ (checked)

# Output:

![Screenshot 2025-05-08 103709](https://github.com/user-attachments/assets/0c5e4b70-4677-428f-a244-ffea53a03c3d)

![Screenshot 2025-05-08 103722](https://github.com/user-attachments/assets/411bbe6b-14c8-40e6-bb8c-bf59d62e9a06)


## Result:
The UiPath workflow successfully reads data from Input.xlsx and writes it to Output.xlsx using Excel activities.
