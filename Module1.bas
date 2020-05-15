Sub CreateWordDocuments()
    Dim CustRow, LastRow As Long
    Dim DocLoc, TagName, FileName As String
    Dim WordDoc, WordApp As Object
    With Sheet1
        DocLoc = ThisWorkbook.Sheets("address").Range("B1") & "\" & ThisWorkbook.Sheets("address").Range("B2")
        On Error Resume Next        'If Word is already Running
        Set WordApp = GetObject("Word.Application")
        If Err.Number <> 0 Then
            'Launch a new instance of Word
            Err.Clear
            'On Error GoTo Error_Handler
            Set WordApp = CreateObject("Word.Application")
            WordApp.Visible = True        'Make the application visible to the user
        End If
        LastRow = .Range("A9999").End(xlUp).Row        'determine the last row in table
        Set WordDoc = WordApp.Documents.Open(FileName:=DocLoc, ReadOnly:=False)
        For CustRow = 2 To LastRow
            Text = .Range("A" & CustRow).Value
            TagName = "(description)"
            With WordDoc.Content.Find
                .Text = TagName
                .Replacement.Text = Text
                .Wrap = wdFindContinue
                .Execute Replace:=wdReplaceOne
            End With
        Next CustRow
        FileName = ThisWorkbook.Path & "\outputword.docx"
        WordDoc.SaveAs FileName
        WordApp.Quit
    End With
End Sub
