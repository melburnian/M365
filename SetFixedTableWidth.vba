Sub SetFixedTableWidth()
    Dim tbl As Table
    
    ' Loop through each table in the active Word document.
    For Each tbl In ActiveDocument.Tables
        ' Set the preferred width using centimetres converted to points.
        tbl.PreferredWidthType = wdPreferredWidthPoints
        tbl.PreferredWidth = Application.CentimetersToPoints(15.98) ' Sets width to 15.98 cm.
        
        ' Disable AutoFit to prevent automatic resizing.
        tbl.AllowAutoFit = False
    Next tbl
End Sub
