' A function for opening up the mac browser using Apply Script
Function BrowseMac(mypath As String) As String
  Dim sMacScript As String
  sMacScript = "set applescript's text item delimiters to "","" " & vbNewLine & _
    "try " & vbNewLine & _
    "set theFiles to (choose file " & _
    "with prompt ""Please select a file or files"" default location alias """ & _
    mypath & """ multiple selections allowed false) as string" & vbNewLine & _
    "set applescript's text item delimiters to """" " & vbNewLine & _
    "on error errStr number errorNumber" & vbNewLine & _
    "return errorNumber " & vbNewLine & _
    "end try " & vbNewLine & _
    "return theFiles"
  BrowseMac = MacScript(sMacScript)
End Function
' Get the basename of a file
Function Basename(pf) As String: Basename = Mid(pf, InStrRev(pf, ":") + 1): End Function
' Check if file exists
Function FileExists(ByVal AFileName As String) As Boolean
    On Error GoTo Catch
    FileSystem.FileLen AFileName
    FileExists = True
    GoTo Finally
Catch:
        FileExists = False
Finally:
End Function
' Extract all references in your data
Sub ExtractReferences(outFile As String)
    Dim cDoc As Word.Document
    Dim cRng As Word.Range
    If FileExists(outFile) Then
        Open outFile For Append As #1
        Set cDoc = ActiveDocument
        Set cRng = cDoc.Content
        cRng.Find.ClearFormatting
        With cRng.Find
            .Forward = True
            .Text = "["
            .Wrap = wdFindStop
            .Execute
            Do While .Found
                cRng.Collapse Word.WdCollapseDirection.wdCollapseEnd
                cRng.MoveEndUntil Cset:="]", Count:=Word.wdForward
                Write #1, cRng.FormattedText
                cRng.Collapse Word.WdCollapseDirection.wdCollapseEnd
                .Execute
            Loop
        End With
        Close #1
    End If
End Sub
' Select fille to write to
Sub SelectFileDialog()
    fullFileName = BrowseMac(MacScript("return (path to documents folder) as String"))
    If fullFileName <> -128 Then
        FileName = Basename(fullFileName)
        ExtractReferences (fullFileName)
        MsgBox "Success! Wrote to " & fullFileName
    End If
End Sub
