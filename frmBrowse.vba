Option Compare Database

Private Sub cmdImport_Click()
    If IsNull(Me.txtFileName) Or Len(Me.txtFileName & "") = 0 Then
    MsgBox "please select excel file"
    Me.cmdSelect.SetFocus
    Exit Sub
End If

Call Initialize
Call ProcessDataWorkbook(Me.txtFileName)

DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel9, "MercImport" & Format(Date, "YYYYMMDD"), Me.txtFileName, True
 
End Sub

Private Sub cmdQuit_Click()
    DoCmd.Close
End Sub

Private Sub cmdSelect_Click()
    Dim strStartDir As String
    
    Dim strFilter As String
    Dim lngFlags As Long
    
    ' Lets start the file browse from our current directory
     
'    strStartDir = CurrentDb.Name
'    strStartDir = Left(strStartDir, Len(strStartDir) - Len(Dir(strStartDir)))
    strStartDir = "R:\Utilization Management\Administration\Vendor Validations\FY18 2017-2018\MERC"
    
    strFilter = ahtAddFilterItem(strFilter, _
                        "Excel Files (*.xls,*.xlsx,*.xlsm)", "*.xls; *.xlsx; *.xlsm")
    Me.txtFileName = ahtCommonFileOpenSave(InitialDir:=strStartDir, _
                     Filter:=strFilter, FilterIndex:=3, Flags:=lngFlags, _
                     DialogTitle:="Select File")
     
End Sub

