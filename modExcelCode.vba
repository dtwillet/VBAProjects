Option Compare Database
Option Explicit
Dim objExcelApp As Object
Dim wb As Object

Sub Initialize()
    Set objExcelApp = CreateObject("Excel.Application")
End Sub
Sub ProcessDataWorkbook(ByVal vWB As Variant, Optional ByVal vWS As Variant = 1)
    Set wb = objExcelApp.Workbooks.Open(vWB)
    Dim ws As Object
    Set ws = wb.Sheets(vWS)
    
    Dim lngLastRow As Long
    Dim lngLastCol As Long
    Dim rng As range
    Dim strRngName As String
        
    strRngName = "MercData"
    lngLastRow = LastRowInColumn()
    lngLastCol = LastColumnInRow()
    With ws
        Set rng = .range(.Cells(1, 1), .Cells(lngLastRow, lngLastCol))
        .Names.Add Name:=strRngName, RefersTo:=rng
        Debug.Print ws.range(strRngName).Address
    End With
    
    wb.Save
    Excel.Application.Quit
    Set wb = Nothing
    
    End Sub

Private Function LastRowInColumn(Optional ByVal vWSName As Variant = "EDIT_LIST", Optional ByVal lngLColumn As Long = 1) As Long
    Dim ws As Worksheet
    Set ws = worksheets(vWSName)
    With ws
        LastRowInColumn = .Cells(.Rows.Count, lngLColumn).End(xlUp).Row
    End With
End Function

Private Function LastColumnInRow(Optional ByVal vWSName As Variant = "EDIT_LIST", Optional ByVal lngLRow As Long = 1) As Long
    Dim ws As Worksheet
    Set ws = worksheets(vWSName)
    With ws
        LastColumnInRow = .Cells(lngLRow, .Columns.Count).End(xlToLeft).Column
    End With
End Function
