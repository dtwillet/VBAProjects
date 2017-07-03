Option Explicit

Function AutoFillNewRecord(F As Form)

   Dim rs As DAO.Recordset, C As Control
   Dim FillFields As String, FillAllFields As Integer
   
   On Error Resume Next
   
   ' Exit if not on the new record.
   If Not F.NewRecord Then Exit Function
   
   ' Goto the last record of the form recordset (to autofill form).
   Set rs = F.RecordsetClone
   rs.MoveLast
   
   ' Exit if you cannot move to the last record (no records).
   If Err <> 0 Then Exit Function
   
   ' Get the list of fields to autofill.
   FillFields = ";" & F![AutoFillNewRecordFields] & ";"
   
   ' If there is no criteria field, then set flag indicating ALL
   ' fields should be autofilled.
   FillAllFields = Err <> 0
   
   F.Painting = False
   
   ' Visit each field on the form.
   For Each C In F
      ' Fill the field if ALL fields are to be filled OR if the
      ' ...ControlSource field can be found in the FillFields list.
      If FillAllFields Or InStr(FillFields, ";" & (C.Name) & ";") > 0 Then
         C = rs(C.ControlSource)
      End If
   Next
   
   F.Painting = True
   
End Function
