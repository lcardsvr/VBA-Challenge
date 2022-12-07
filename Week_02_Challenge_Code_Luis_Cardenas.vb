Sub Stock_Analysis()


Dim i As LongLong           'Counter to go through all the individual ticker values in the sheet
Dim j As LongLong           'Counter to document summary results for Ticker, Yearly Change, Percent Change and Total Stock Volume
Dim k As Integer            'Counter with number of Sheets in the document
Dim Open_Value As Double    'Open_Value per ticker
Dim Close_Value As Double   'Close_Value per ticker
Dim Total_Volume As LongLong 'Total volume per ticker
Dim Ticker As String        'Ticker string
Dim Yearly_Change As Double 'Yearly change per ticker
Dim Per_Change As Double    'Percentage change per ticker
Dim Num_Worksheets As Integer   'Total number of sheets
Dim pos_great_increase As Integer 'Greatest Increase row position in summary data
Dim pos_great_decrease As Integer  'Greatest decrease row position in summary data
Dim pos_great_volume As Integer     ''Greatest Volume row position in summary data

'Number of worksheets

Num_Worksheets = ThisWorkbook.Sheets.Count


'Cycle to go over all the sheets in the document

For k = 1 To Num_Worksheets

    Worksheets(1).Activate

    Sheets(k).Select
    
    i = 2

    j = 1

    'Labels the summary titles
    
    Cells(j, 9).Value = "Ticker"
    
    Cells(j, 10).Value = "Yearly Change"
    
    Cells(j, 11).Value = "Percent Change"
    
    Cells(j, 12).Value = "Total Stock Volume"
    
    j = 2

    'Cycle to go through the sheet until an empty cell is found
    
    Do While Cells(i, 1).Value <> Empty


        Ticker = Cells(i, 1).Value
        
        Open_Value = Cells(i, 3).Value  'First day of a given ticker. Open value is recorded
        
        Total_Volume = 0    'Starting value for the volume to be cumulatively calculated

        
        'Cycle to go through the sheet while the ticker is the same
        
        Do While Cells(i, 1).Value = Cells(i + 1, 1).Value

            'Total volume is cumulatively calculated
            
            Total_Volume = Total_Volume + Cells(i, 7).Value
            
            i = i + 1

        Loop
    
        
        'i contains the information of the last day of the year for a given ticker
        
        'Last cell for Total Volume added to the cumulatively calculated value
        
        Total_Volume = Total_Volume + Cells(i, 7).Value
    
        
        'Lastday of a given ticker. Close value is recorded
        
        Close_Value = Cells(i, 6).Value
    
        
        'Calculated values for a Ticker
        
        Yearly_Change = Close_Value - Open_Value
    
        Per_Change = Close_Value / Open_Value - 1
    
    
        'Calculated values recording in spreadsheet
        
        Cells(j, 9).Value = Ticker
    
        Cells(j, 10).Value = Yearly_Change
        
        Cells(j, 11).Value = Per_Change
    
        Cells(j, 12).Value = Total_Volume
    
        i = i + 1
    
        j = j + 1
    
    Loop
    
    'After all the values for Ticker, Yearly Change, Percent Change and Total Stock Volume
    'are available in the spreadsheet, the row for the greatest volume, greatest increase
    'and decrease are obtained. Once identified, the ticker and respective values are
    'recorded in the spreadshee in the nominated cell ranges.
    
    
    pos_great_volume = Greatest_Volume()

    pos_great_increase = Greatest_Increase()

    pos_great_decrease = Greatest_Decrease()



    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"

    Range("O2").Value = "Greatest % Increase"
    Range("P2").Value = Cells(pos_great_increase, 9).Value
    Range("Q2").Value = Cells(pos_great_increase, 11).Value

    Range("O3").Value = "Greatest % Decrease"
    Range("P3").Value = Cells(pos_great_decrease, 9).Value
    Range("Q3").Value = Cells(pos_great_decrease, 11).Value


    Range("O4").Value = "Greatest Total Volume"
    Range("P4").Value = Cells(pos_great_volume, 9).Value
    Range("Q4").Value = Cells(pos_great_volume, 12).Value


    'Once all the values are in their respective position, the conditional formatting for the
    'yearly changes, the percentage and spacing are adjusted in the FixFormatting function.
    
    Fix_Formatting


Next k

'Once all the sheet work has taken place, the first sheet is selected again

Sheets(1).Select

End Sub


'------ Function to obtain the Greatest Volume
Function Greatest_Volume() As Integer


Dim i As Integer
Dim Io As Double
Dim j As Integer


Io = Cells(2, 12).Value

j = 2
i = 2

Do While Cells(i, 9).Value <> Empty

    If Cells(i, 12).Value > Io Then
    
        Io = Cells(i, 12).Value
        j = i
    
    End If
    
    i = i + 1


Loop

Greatest_Volume = j

End Function
'------ Function to obtain the Greatest Increase
Function Greatest_Increase() As Integer


Dim i As Integer
Dim Io As Double
Dim j As Integer


Io = Cells(2, 11).Value

j = 2
i = 2

Do While Cells(i, 9).Value <> Empty

    If Cells(i, 11).Value > Io Then
    
        Io = Cells(i, 11).Value
        j = i
    
    End If
    
    i = i + 1


Loop

Greatest_Increase = j

End Function

'------ Function to obtain the Greatest Decrease

Function Greatest_Decrease() As Integer


Dim i As Integer
Dim Io As Double
Dim j As Integer


Io = Cells(2, 11).Value

j = 2
i = 2

Do While Cells(i, 9).Value <> Empty


    If Cells(i, 11).Value < Io Then
    
        Io = Cells(i, 11).Value
        j = i
    
    End If
    
    i = i + 1


Loop

Greatest_Decrease = j


End Function

'------ Function to ajdust Color Change on values, percentage and cell spacing formatting

Function Fix_Formatting()


    
'Change Yearly - Change Color Format

    Columns("J:J").Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreaterEqual _
        , Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 10092441
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Cells.FormatConditions.Delete
    Columns("J:J").Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Columns("J:J").Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreaterEqual _
        , Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 10092441
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    
 'Remove color from title
 
    Range("K1").Select
    Selection.Copy
    Range("J1").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False

'Fix percentage format

    Range("K:K,Q2,Q3").Select
    Selection.NumberFormat = "0.00%"
    
    'Fix column size
    Cells.Select
    Cells.EntireColumn.AutoFit
    
    Range("A1").Select

End Function



