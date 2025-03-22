Sub RemoveExcessLineSpacing()
    Dim i As Long          ' Loop counter for paragraphs.
    Dim emptyCount As Long ' Counter for consecutive empty paragraphs.
    Dim paraText As String ' Holds the text of the paragraph.
    
    emptyCount = 0
    
    ' Loop backwards from the last paragraph to the first.
    For i = ActiveDocument.Paragraphs.Count To 1 Step -1
        ' Get the text of the current paragraph.
        paraText = ActiveDocument.Paragraphs(i).Range.Text
        ' Remove carriage returns and line feeds for a proper check.
        paraText = Trim(Replace(paraText, vbCr, ""))
        paraText = Trim(Replace(paraText, vbLf, ""))
        
        ' Check if the paragraph is empty.
        If paraText = "" Then
            emptyCount = emptyCount + 1
        Else
            ' If more than two empty paragraphs were found consecutively,
            ' delete extra ones leaving only one.
            If emptyCount > 2 Then
                Dim j As Long
                ' Delete extra blank paragraphs (emptyCount - 1), leaving one.
                For j = 1 To emptyCount - 1
                    ' Delete the blank paragraph immediately following the current non-empty paragraph.
                    ActiveDocument.Paragraphs(i + 1).Range.Delete
                Next j
            End If
            ' Reset the counter when a non-empty paragraph is encountered.
            emptyCount = 0
        End If
    Next i
    
    ' Handle the case where the document begins with more than two blank paragraphs.
    If emptyCount > 2 Then
        Dim k As Long
        For k = 1 To emptyCount - 1
            ActiveDocument.Paragraphs(1).Range.Delete
        Next k
    End If
End Sub
