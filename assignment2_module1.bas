Attribute VB_Name = "Module1"
Option Explicit

' ==== CP212 Windows Application Programming ===============+
' Name: Connor Logan
' Student ID: 190209360
' Date: 2020-10-06
' Program title: Assignment 2
' Description: Refer to the description in the assignment 2 PDF
'===========================================================+
Sub worksheetCreation()

'Variable Declaration

Dim columnRange As Range
Dim columnCell As Range
Dim wsBool As Boolean
Dim Sheet As Worksheet
Dim worksheetBool As Boolean
Dim i As Integer
Dim productCell As Range

'Sets boolean to true

wsBool = True

'Sets the range to be the entire category column
Set columnRange = Range("B4", Range("B4").End(xlDown))

'Nested for each loop
'That walks through the cells and sheets
'Creates new worksheets based on product type
'ensures no duplicates with the boolean

For Each columnCell In columnRange
    For Each Sheet In ActiveWorkbook.Worksheets
        If columnCell.Value = Sheet.Name Then
            wsBool = False
            
            'This block of code
            'Formats the cell with width,
            'And everything gets centered
            'Inserts the Products and prices
            'Per Category
            Range("A3").ColumnWidth = 45
            Range("A3").HorizontalAlignment = xlCenter
            Range("B3").ColumnWidth = 45
            Range("B3").HorizontalAlignment = xlCenter
            Range("A3").Value = "Products in " & columnCell.Value & " Category"
            Range("B3").Value = "Prices in " & columnCell.Value & " Category"
           
            ' The idea here is so ensure that I am grabbing
            ' Data that matches with the sheet name
            ' And insert product names
            
            i = 0
            For Each productCell In columnRange
                
                If productCell.Value = Sheet.Name Then
                    
                    Range("A4").Offset(i, 0).Value = productCell.Offset(0, -1).Value
                    Range("B4").Offset(i, 0).Value = productCell.Offset(0, 1).Value
                    i = i + 1
                End If
                
            Next productCell
            
        End If
        
        
    Next Sheet
    'Creates new sheet if the name doesn't exist
    If wsBool = True Then
        ActiveWorkbook.Worksheets.Add.Name = columnCell.Value
    End If
    wsBool = True
Next columnCell

End Sub
