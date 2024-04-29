# BikeStoreDashboard
**Pet-project that contains creating dashboard for Bikes Store with SQL and Excel**  
bike-store-analysis.ipynb File with SQL queries  
result.csv File that i got to edit and use with Excel  
Dashboard.xlsx Final file with Dashboard  
Dashboard.xlsm Final file that contains VBA macros  

# SQL part of the project
1. Get data from relational database and create connection with database
2. Create SQL query to get number of completed orders by store
3. Create SQL query using the CTE to get percentage of completed orders by store
4. Create SQL query to get total customers who ordered products from store
5. Create SQL query using subquery to get clients whose total purchases are above average
6. Create wide table including the necessary columns from the database for further analysis in Excel
7. Save data as csv-file


# Excel part of the project
1. Get data to Excel from csv-file
2. Change data types (Change comma to period in decimal numeric data types and round them) using Excel Power Query   
```
= Table.TransformColumns(#"Измененный тип2",{{"total_spent", each Number.Round(_, 2), type number}})   
--> for list_price/discount/total_spent
```
3. Create pivot tables for out future pivot diagrams and dashboard
Top cities by total income, Top brands, Top produsct, Top categories, Total income by state and Income by months
4. Create different types of diagrams. Edit and custimize them
5. Create slicer and connect it to each diagram
6. Create and custimize interactive (by slicers) dashboard

7. Create feature by using Excel macroses and VBA language:
After entering the customer's name and pressing the button, Their order_id and the amount of money they spent is automatically calculated and displayed on the message box

Below is the VBA code for the macros, dashboard and macros using
 
 
![photo_2024-04-25_18-37-08](https://github.com/Yurii-Molotow/BikeStoreDashboard/assets/168109152/d1882ea8-95df-4e72-819e-65a47c86bd40)
![photo_2024-04-25_18-37-15](https://github.com/Yurii-Molotow/BikeStoreDashboard/assets/168109152/c45da616-e414-4854-9042-d19dde4cfe09)




```
Option Explicit

Sub Кнопка_1()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim nameCol As Long, orderIDCol As Long, amountCol As Long
    Dim clientName As String, orderID As String
    Dim totalAmount As Double
    Dim i As Long
    
    Set ws = ThisWorkbook.Sheets("Äàí³")
    
    lastRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
    
    clientName = ws.Range("B2").Value
    
    nameCol = ws.Rows(1).Find("customer_name").Column
    orderIDCol = ws.Rows(1).Find("order_id").Column
    amountCol = ws.Rows(1).Find("total_spent").Column
    
    For i = 2 To lastRow
        If ws.Cells(i, nameCol).Value = clientName Then
            orderID = ws.Cells(i, orderIDCol).Value
            totalAmount = 0
            
            Do While ws.Cells(i, orderIDCol).Value = orderID
                totalAmount = totalAmount + ws.Cells(i, amountCol).Value
                i = i + 1
                
                If i > lastRow Then Exit Do
            Loop
            
            i = i - 1

            MsgBox "Ім'я: " & clientName & vbCrLf & _
                   "Номер замовлення: " & orderID & vbCrLf & _
                   "Загальна вартість: " & totalAmount

        End If
    Next i

End Sub

```
