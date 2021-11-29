Private Sub Worksheet_SelectionChange(ByVal target As Range)

If Selection.Count = 1 Then
'--------------------------------------------------------------------

'H:H triggers IssueWorkbookLooper
    If InStr(1, (Range("AA5").Value), "On") > 0 Then
        If Not Intersect(target, Range("$H:$H")) Is Nothing Then
            Call IssueSheetGen
        End If
    End If
'--------------------------------------------------------------------

'S1 triggers MacrosOn
    If Not Intersect(target, Range("$S$1")) Is Nothing Then
        Call macrosOn
    End If

'T1 triggers MacrosOff
    If Not Intersect(target, Range("$T$1")) Is Nothing Then
        Call macrosOff
    End If

End If

End Sub
