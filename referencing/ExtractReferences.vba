Sub ExtractReferences()
    Dim cDoc As Word.Document
    Dim cRng As Word.Range
    Dim outFile As String
    
    outFile = "/Users/heymann/Desktop/test/TestResults.txt"
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
End Sub
