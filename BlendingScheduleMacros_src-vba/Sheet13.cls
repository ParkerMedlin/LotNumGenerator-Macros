Private Sub Worksheet_SelectionChange(ByVal target As Range)

    If InStr(1, (Range("AA4").Value), "On") > 0 Then
        If Selection.Count = 1 Then
            If Not Intersect(target, Range("$R:$R")) Is Nothing Then
                Call LotNumGen
            End If
        End If
    End If

End Sub
