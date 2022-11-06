Public Sub Add_Images_To_Cells()

    Dim lastRow As Long
    Dim URLs As Range, URL As Range
    Dim pic As Picture
    Dim urlColumn As String
    
    With ActiveSheet
        urlColumn = "D"
        lastRow = .Cells(Rows.Count, urlColumn).End(xlUp).Row
        Set URLs = .Range(urlColumn & "2:" & urlColumn & lastRow)
    End With

    For Each URL In URLs
        If InStr(URL.Value, "http") > 0 Then
            URL.Offset(0, 0).Select
            Set pic = URL.Parent.Pictures.Insert(URL.Value)
            With pic.ShapeRange
                .LockAspectRatio = msoFalse
                .Height = URL.Offset(0, 0).Height - 1
                .Width = URL.Offset(0, 0).Width - 1
                .LockAspectRatio = msoTrue
            End With
            
            URL.Clear
            DoEvents
        End If
    Next
    
End Sub