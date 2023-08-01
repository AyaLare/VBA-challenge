Attribute VB_Name = "Module1"
Sub Stock_Display()
'Declaring the primary variables
    Dim wb As Workbook
    Dim NumSheets As Long
    Dim sheet As Excel.Worksheet
    Dim Num_Ticker_Symbols As Long
    Dim Ticker_Symbol As String
    Dim Opening_Price As Double
    Dim Closing_Price As Double
    Dim Stock_Date As String
    Dim Price_Year_Change As Double
    Dim Price_Change_Percent As Double
    Dim Total_Stock_Volume As Double
    Dim Rownum As Long
    Dim iRange As Range
'Declaring variables to get the greatest figures
    Dim Greatest_Percent_Increase As Double
    Dim Greatest_Percent_Increase_Symbol As String
    Dim Greatest_Percent_Decrease As Double
    Dim Greatest_Percent_Decrease_Symbol As String
    Dim Greatest_Total_Volume As Double
    Dim Greatest_Total_Volume_Symbol As String
'
'Determine the number of sheets
'
NumSheets = ThisWorkbook.Worksheets.Count
MsgBox ("Number of Sheets " + CStr(NumSheets))
For j = 1 To NumSheets
' Set the code to point to the Sheet
Set sheet = ThisWorkbook.Sheets(j)
MsgBox (sheet.Name)
'
'Write the headers for the data to be analysed
'
sheet.Range("I1") = "Ticker"
sheet.Range("J1") = "Yearly change"
sheet.Range("K1") = "Percent Change"
sheet.Range("L1") = "Total Stock Volume"
'
'starting the logic to identify the Ticker Symbol
'
'Set the data variable to their starting point for each sheet
'
    Rownum = 0
    Num_Ticker_Symbols = 0
    Ticker_Symbol = ""
    Total_Stock_Volume = 0
    Greatest_Percent_Increase = 0
    Greatest_Percent_Increase_Symbol = ""
    Greatest_Percent_Decrease = 0
    Greatest_Percent_Decrease_Symbol = ""
    Greatest_Total_Volume = 0
    Greatest_Toltal_Volume_Symbol = ""
    Price_Year_Change = 0

' Start the loop through the entire data
' Get the number of rows
    Rownum = sheet.Cells(Rows.Count, 1).End(xlUp).Row
      
    ' MsgBox ("Number of Rows " + CStr(Rownum))
     ' For i = 2 To 75301
     For i = 2 To Rownum
    'To display the Row Number
    'MsgBox ("Row Number is " + CStr(i))
        '
        'Testing for a change of Ticker Symbol
        If Ticker_Symbol = Cells(i, 1) Then
        '
        'Increase the stock volume by the volume of the day
            Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7)
        
        'MsgBox ("Ticker Symbol has not changed")
        
            Else
            'Ticker Symbol has changed
                'MsgBox ("Ticker symbol has Changed to " + Cells(i, 1))
                Num_Ticker_Symbols = Num_Ticker_Symbols + 1
                Ticker_Symbol = Cells(i, 1)
                Total_Stock_Volume = Cells(i, 7)
                     
        End If
        'Now checking for the start of a new year
        'MsgBox ("Month and Day is :" + Right(Cells(i, 2), 4))
        If Right(Cells(i, 2), 4) = "0102" Then
            Opening_Price = Cells(i, 3)
            'MsgBox ("Ticker: " + Ticker_Symbol + " " + "Year Opening Price: " + CStr(Opening_Price))
        End If
                     
        'Now checking for the end of a new year
        'MsgBox ("Month and Day is :" + Right(Cells(i, 2), 4))
        If Right(Cells(i, 2), 4) = "1231" Then
            Closing_Price = Cells(i, 6)
            Price_Year_Change = Closing_Price - Opening_Price
            Price_Change_Percent = Round((((Price_Year_Change) / Opening_Price) * 100), 2)
                 'MsgBox ("Ticker: " + Ticker_Symbol + " " + "Year Opening Price: " + CStr(Opening_Price) + " " + "Year Closing Price: " + CStr(Closing_Price) + "  Percentage change from the opening price is: " + CStr(Price_Change_Percent) + "%" + " Total Stock Volume is: " + CStr(Total_Stock_Volume))
                'Write the result for the stock in the sheet
                '
                ' Write the ticker symbol
                '
                sheet.Cells(Num_Ticker_Symbols + 1, 9).Value = CStr(Ticker_Symbol)
                '
                'Write the yearly price change
                '
                sheet.Cells(Num_Ticker_Symbols + 1, 10).Value = CStr(Price_Year_Change)
                '
                ' Conditionally color the Yearly Price Change cells
                ' Reference to color index https://www.excel-easy.com/vba/examples/background-colors.html
                '
                If Price_Year_Change >= 0 Then
                    sheet.Cells(Num_Ticker_Symbols + 1, 10).Interior.ColorIndex = 3
                    Else
                     sheet.Cells(Num_Ticker_Symbols + 1, 10).Interior.ColorIndex = 4
                End If
                '
                 'Write the percent change price
                 '
                sheet.Cells(Num_Ticker_Symbols + 1, 11).Value = CStr(Price_Change_Percent)
                '
                'Conditionally color the Yearly Price Percent Change cells
                '
                If Price_Change_Percent >= 0 Then
                    sheet.Cells(Num_Ticker_Symbols + 1, 11).Interior.ColorIndex = 3
                    Else
                     sheet.Cells(Num_Ticker_Symbols + 1, 11).Interior.ColorIndex = 4
                End If
                '
                'Write the total stock volume
                '
                sheet.Cells(Num_Ticker_Symbols + 1, 12).Value = CStr(Total_Stock_Volume)
 
        End If
        '
        ' Test for the greatest stock volume
        If Total_Stock_Volume > Greatest_Total_Volume Then
                Greatest_Total_Volume = Total_Stock_Volume
                Greatest_Total_Volume_Symbol = Ticker_Symbol
            End If
         'Test for the greatest price increase percent
            If Greatest_Percent_Increase < Price_Change_Percent Then
                Greatest_Percent_Increase = Price_Change_Percent
                Greatest_Percent_Increase_Symbol = Ticker_Symbol
            End If
            ' Test for the greatest price decrease percent
            If Greatest_Percent_Decrease > Price_Change_Percent Then
                Greatest_Percent_Decrease = Price_Change_Percent
                Greatest_Percent_Decrease_Symbol = Ticker_Symbol
                'MsgBox ("Current Greatest Percent Decrease" + Greatest_Price_Decerease_Symbol + " " + CStr(Greatest_Percent_Decrease) + "%")
            End If
        
    Next i
    'Write the headers and the year summary analysis
    '
        sheet.Range("P1") = "Ticker"
        sheet.Range("Q1") = "Value"
        sheet.Range("O2") = "Greatest % Increase: "
        sheet.Range("P2") = Greatest_Percent_Increase_Symbol
        sheet.Range("Q2") = CStr(Greatest_Percent_Increase) + "%"
        sheet.Range("O3") = "Greatest % Decrease: "
        sheet.Range("P3") = Greatest_Percent_Decrease_Symbol
        sheet.Range("Q3") = CStr(Greatest_Percent_Decrease) + "%"
        sheet.Range("O4") = "Greatest Total Volume: "
        sheet.Range("P4") = Greatest_Total_Volume_Symbol
        sheet.Range("Q4") = CStr(Greatest_Total_Volume)
    'MsgBox ("Greatest % Increase: " + Greatest_Percent_Increase_Symbol + " " + CStr(Greatest_Percent_Increase) + "%")
    'MsgBox ("Greatest % Decrease: " + Greatest_Percent_Decrease_Symbol + " " + CStr(Greatest_Percent_Decrease) + "%")
    'MsgBox ("Greatest Total Volume: " + Greatest_Total_Volume_Symbol + " " + CStr(Greatest_Total_Volume))
    
    Next j
    
End Sub

