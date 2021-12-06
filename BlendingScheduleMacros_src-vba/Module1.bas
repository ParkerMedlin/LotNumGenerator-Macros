Sub Macro1()
'
' Macro1 Macro
'

'
    ActiveSheet.ListObjects("blendData").Range.AutoFilter Field:=2, Criteria1:= _
        Array("14308.B", "14308AMBER.B", "93100DSL.B", "93100GAS.B", "93100TANK.B"), _
        Operator:=xlFilterValues
End Sub