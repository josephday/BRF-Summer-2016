Option Compare Text

Private Sub CheckBox1_Click()

End Sub

Private Sub CheckBox2_Click()

End Sub

Private Sub CheckBox4_Click()

End Sub

Private Sub CommandButton1_Click()
    
    sbADOExample
    
End Sub


Sub sbADOExample()
Dim sSQLQry As String
Dim ReturnArray
Dim Conn As New ADODB.Connection
Dim mrs As New ADODB.Recordset
Dim DBPath As String, sconnect As String
Dim strMystring As String
Dim runningtotal As Integer
runningtotal = 0
strMystring = "SELECT * From [data$] "
DBPath = ThisWorkbook.FullName
'You can provide the full path of your external file as shown below
'DBPath ="C:\InputData.xlsx"
sconnect = "Provider=MSDASQL.1;DSN=Excel Files;DBQ=" & DBPath & ";HDR=Yes';"
Conn.Open sconnect

If Sheets("View").OLEObjects("Checkbox1").Object.Value = True Then
strMystring = AddClientQuery(strMystring)
runningtotal = runningtotal + 1
End If

If Sheets("View").OLEObjects("Checkbox2").Object.Value = True Then
strMystring = AddProviderQuery(strMystring, runningtotal)
runningtotal = runningtotal + 1
End If

If Sheets("View").OLEObjects("Checkbox3").Object.Value = True Then
strMystring = AddFundQuery(strMystring, runningtotal)
runningtotal = runningtotal + 1
End If

If Sheets("View").OLEObjects("Checkbox4").Object.Value = True Then
strMystring = AddMinComQuery(strMystring, runningtotal)
runningtotal = runningtotal + 1
End If

If Sheets("View").OLEObjects("Checkbox5").Object.Value = True Then
strMystring = AddMaxComQuery(strMystring, runningtotal)
runningtotal = runningtotal + 1
End If


'MsgBox (strMystring)

  ' Your SQL Statement (Table Name= Sheet Name=[Sheet1$])
mrs.Open strMystring, Conn
'=>Load the Data into an array
'ReturnArray = mrs.GetRows
''OR''
'=>Paste the data into a sheet
With Sheets("View")
    .Rows("20" & ":" & .Rows.Count).ClearContents
End With
Sheet1.Range("A20").CopyFromRecordset mrs
'Close Recordset
mrs.Close
'Close Connection
Conn.Close
End Sub

Function AddClientQuery(query As String)
Dim add1 As String
add1 = InputBox("Enter Client's Surname: ", "Enter Name")

AddClientQuery = query & _
"WHERE [Client] = " & Chr$(39) & UCase(add1) & Chr$(39)
End Function

Function AddProviderQuery(query As String, total As Integer)
Dim add2 As String
add2 = InputBox("Enter Provider: ", "Enter Name")

If total = 0 = True Then
AddProviderQuery = query & _
"WHERE [Provider] = " & Chr$(39) & UCase(add2) & Chr$(39)
ElseIf total <> 0 = True Then
AddProviderQuery = query & _
" AND [Provider] = " & Chr$(39) & UCase(add2) & Chr$(39)
End If
End Function

Function AddFundQuery(query As String, total As Integer)
Dim add3 As String
add3 = InputBox("Enter Fund: ", "Enter Name")

If total = 0 = True Then
AddFundQuery = query & _
"WHERE [Fund] = " & Chr$(39) & UCase(add3) & Chr$(39)
ElseIf total <> 0 = True Then
AddFundQuery = query & _
" AND [Fund] = " & Chr$(39) & UCase(add3) & Chr$(39)
End If
End Function

Function AddMinComQuery(query As String, total As Integer)
Dim add4 As String
add4 = InputBox("Enter Minimum Commitment (No commas): ", "Enter Number")
If total = 0 = True Then
AddMinComQuery = query & _
"WHERE [Commitment] >= " & add4
ElseIf total <> 0 = True Then
AddMinComQuery = query & _
" AND [Commitment] >= " & add4
End If
End Function

Function AddMaxComQuery(query As String, total As Integer)
Dim add5 As String
add5 = InputBox("Enter Maximum Commitment (No commas) : ", "Enter Number")
If total = 0 = True Then
AddMaxComQuery = query & _
"WHERE [Commitment] <= " & add5
ElseIf total <> 0 = True Then
AddMaxComQuery = query & _
" AND [Commitment] <= " & add5
End If
End Function



Private Sub CommandButton2_Click()
    CopyFromWorksheets
End Sub
