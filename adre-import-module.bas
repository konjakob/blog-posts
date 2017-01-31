Attribute VB_Name = "Module1"



Sub createSheets()
Attribute createSheets.VB_ProcData.VB_Invoke_Func = " \n14"

    Dim names(1 To 6) As String
    Dim val(1 To 6) As String
    Dim n As Integer
    
    names(1) = "FLOW"
    names(2) = "PS1"
    names(3) = "PS2"
    names(4) = "PS3"
    names(5) = "TORQUE"
    names(6) = "PS0"
    
    val(1) = "Flow"
    val(2) = "Pressure 1"
    val(3) = "Pressure 2"
    val(4) = "Pressure 3"
    val(5) = "Torque"
    val(6) = "Pressure 0"
    
    ' first create for each channel a single worksheet
    For n = 1 To UBound(names)
        Sheets.Add After:=Sheets(Sheets.Count)
        Sheets(Sheets.Count).name = val(n)
    Next n
    
    r = Sheets("Sheet1").Cells(Rows.Count, 2).End(xlUp).Row 'find last row
 
    ' add filter for each channel and copy the date to the respective sheet
    Sheets("Sheet1").Select
    Range("A3:H" & r).AutoFilter
    
    For n = 1 To UBound(names)
        Sheets("Sheet1").Select
        Range("A3:H" & r).AutoFilter Field:=1, Criteria1:=Application.WorksheetFunction.Index(Sheets("Sheet1").Range("A:A"), Application.WorksheetFunction.Match(names(n), Sheets("Sheet1").Range("B:B"), 0))
        Range("A3:H" & r).AutoFilter Field:=3, Criteria1:="<>Pump", Operator:=xlAnd
        Range("A3:H" & r).AutoFilter Field:=6, Criteria1:="0", Operator:=xlAnd
        
        Worksheets("Sheet1").Range(Cells(3, 1), Cells(r, 8)).SpecialCells(xlCellTypeVisible).Copy
        Sheets(val(n)).Select
        Worksheets(val(n)).Paste
    Next n
   
    ' remove the view from the filter
    Sheets("Sheet1").Select
    Range("A3:H" & r).AutoFilter Field:=1
    
   ' format the date column to time
   For n = 1 To UBound(names)
    Sheets(val(n)).Select
    Columns("E:E").Select
    Selection.NumberFormat = "dd/mm/aaaa h:mm:ss"
    
    Columns("H:H").Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=NOT(ISNUMBER(H1))"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
    End With
    
   Next n
   
End Sub
