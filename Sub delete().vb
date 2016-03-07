Sub delete()
    'Worksheets("mapping").Range("F3:Z17").ClearContents
    Sheets("mapping").Select
    
        Dim S As Shape
            Dim RG As Range
            For Each S In Sheets("mapping").Shapes
                If S.Type <> 8 Then
                    S.delete
                End If
            Next S
End Sub

