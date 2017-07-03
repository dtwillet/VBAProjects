Option Compare Database
Option Explicit
Sub BrowseFileToImport()
    DoCmd.OpenForm "frmBrowse"
End Sub

Sub UpdateQuery(strShtImport As String)
    Dim db As DAO.Database
    Dim td As DAO.TableDef
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim lBatchID As Long
    Dim iCount As Integer
    
    
    Set db = CurrentDb()
    
    'hardcoded sheet
    'strShtImport = "R:\Utilization Management\Administration\Vendor Validations\FY18 2017-2018\MERC\APRIL 2017\April 2017.xlsx"
    
    sql = "INSERT INTO tblBatch ( BatchDate, VENDOR_NAME, VendorID, Source )" _
            & " VALUES ('" & Date & "' , 'MERC', 4, '" & strShtImport & "');"

    db.Execute sql, dbFailOnError
    iCount = db.RecordsAffected
    Debug.Print sql & vbCr & iCount

    'Find Latest BatchID
    lBatchID = DMax("BatchID", "[tblBatch]")
    lBatchID = 5
    
    On Error Resume Next
    
    'Delete temp table invoiceable records
    sql = "DELETE * FROM tblTempMercOrderInvoiceable"
    
    db.Execute sql, dbFailOnError
    iCount = db.RecordsAffected
    Debug.Print sql & vbCr & iCount
    
    'insert from query into temp table invoiceable
    sql = "INSERT INTO tblTempMercOrderInvoiceable ( PATIENT_ID, ORDER_ID, CPT_CODE_ID, ITEM_STATUS, ORDER_STATUS ) " _
            & "SELECT qryMercOrderInvoiceable.PATIENT_ID, qryMercOrderInvoiceable.ORDER_ID, qryMercOrderInvoiceable.CPT_CODE_ID, " _
            & "qryMercOrderInvoiceable.ITEM_STATUS, qryMercOrderInvoiceable.ORDER_STATUS FROM qryMercOrderInvoiceable;"
    
    db.Execute sql, dbFailOnError
    iCount = db.RecordsAffected
    Debug.Print sql & vbCr & iCount
    
    'if temp import table exists then drop
    For Each td In db.TableDefs
        If td.Name = "tblTempMercImport" Then
            db.Execute "Drop Table tblTempMercImport;"
        End If
    Next
    
    'import spreadsheet to temp table
    DoCmd.TransferSpreadsheet acImport, , "tblTempMercImport", strShtImport, True, "CleanInvoices!ExternalData_1"
    
    'create recordset to count records in temp table
    Set rs = db.OpenRecordset("select * from tblTempMercImport")
    rs.MoveLast
    iCount = rs.RecordCount
    Debug.Print "Spreadsheet Import " & vbCr & iCount
    
     'Import temp table to tblMercImport wtih BatchID and BatchLine
      sql = "INSERT INTO tblMercImport ( InvoiceNbr, PatientName, PatientDOB, PatientPhysician, PatientAuthEnd, PatientDX, PatientMRN, ServiceDate, ProcCode, Modifier, Units, ChargeAmt, Comments, CleanMRN, CPTBase, CPTFull, MstCPTIDMatch, VendorName, BatchID, PatientID )" _
            & " SELECT tblTempMercImport.InvoiceNbr, tblTempMercImport.PatientName, tblTempMercImport.PatientDOB, tblTempMercImport.PatientPhysician, tblTempMercImport.PatientAuthEnd, tblTempMercImport.PatientDX, tblTempMercImport.PatientMRN, tblTempMercImport.ServiceDate, tblTempMercImport.ProcCode, tblTempMercImport.Modifier, tblTempMercImport.Units, tblTempMercImport.ChargeAmt, tblTempMercImport.Comments, tblTempMercImport.CleanMRN, tblTempMercImport.CPTBase, tblTempMercImport.CPTFull, tblTempMercImport.MstCPTIDMatch, tblTempMercImport.VendorName, '" & lBatchID & "', tblTempMercImport.PatientID " _
            & "FROM tblTempMercImport;"
    
    db.Execute sql, dbFailOnError
    iCount = db.RecordsAffected
    Debug.Print sql & vbCr & iCount

   
    'Update tblMercImport with OrderID from tblTempMercOrderInvoiceable
    sql = "UPDATE tblMercImport INNER JOIN tblTempMercOrderInvoiceable ON " _
            & "(tblMercImport.CPTFull = tblTempMercOrderInvoiceable.CPT_CODE_ID) AND " _
            & "(tblMercImport.PatientID = tblTempMercOrderInvoiceable.PATIENT_ID) " _
            & "SET tblMercImport.OrderID = [tblTempMercOrderInvoiceable].[ORDER_ID], " _
            & "tblMercImport.OrderStatus = [tblTempMercOrderInvoiceable].[ITEM_STATUS], " _
            & "tblMercImport.FinalPayment = [tblTempMercOrderInvoiceable].[ORDER_STATUS]" _
            & " WHERE (((tblMercImport.BatchID)=" & lBatchID & "));"

    db.Execute sql, dbFailOnError
    iCount = db.RecordsAffected
    Debug.Print sql & vbCr & iCount
    
End Sub


