Attribute VB_Name = "mdlGeneral"
Public Function GetSERP(Query As String, ApIKey As String) As String
  Dim sheaders As String
  Dim scontent As String
  Dim Json As Object
  Dim companyWebsite As String
  
  With CreateObject("MSXML2.ServerXMLHTTP")
      .Open "GET", "https://app.zenserp.com/api/v2/search?q=" & WorksheetFunction.EncodeURL(Query) & "&location=United+States&search_engine=google.com", False
      .SetRequestHeader "apikey", ApIKey
      .Send
      sheaders = .getAllResponseHeaders
      scontent = .responseText
  End With

  On Error GoTo errHandler
  GetSERP = scontent
Exit Function

errHandler:
    MsgBox "API returned the following JSON error message: " & scontent

End Function

Sub RunAPIQuery()
    If MsgBox("Do you want to run the API query with the values provided in column A?", vbYesNo, "Confirm") = vbYes Then
        Dim CompanyName As String
        Dim myAPIKey As String
        Dim ws As Worksheet
        Dim lastRow As Double
        
        Set ws = ActiveWorkbook.ActiveSheet
        
        With ws
            myAPIKey = .Cells(1, 5)
            lastRow = .Cells(1, 1).End(xlDown).Row
            If lastRow > 1000000# Then
                MsgBox "No data for query.", vbCritical, "Error"
                Exit Sub
            End If
            For i = 2 To lastRow
                CompanyName = .Cells(i, 1)
                .Cells(i, 2) = GetSERP(CompanyName, myAPIKey)
            Next i
        End With
        MsgBox "Done!", vbInformation, "Done!"
    Else
        Exit Sub
    End If
End Sub


