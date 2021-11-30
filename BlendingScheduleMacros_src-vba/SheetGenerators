Sub blndShtGen3wk()
'///Assigned to the 3wk column on Blend These//////////////////////////////////////////////////////////////////////////////
'///Opens the blend sheet workbook for the blend on the row you click. Inputs the 3wk qty into said blend sheet.///////////

'String for current workbook name
    Dim src As String
    src = ActiveWorkbook.Name

'Copy the value of the shortage qty cell
    Selection.Copy
    
'Get the filepath + name of the appropriate blend sheet
    Dim blndShtPath As String
    blndShtPath = ActiveCell.Offset(0, -3).Value
    Workbooks.Open Filename:= _
        blndShtPath
    
'Paste the shortage qty value into the theory gallons cell
    Cells.Find(What:="Theory Gallons", After:=ActiveCell, LookIn:=xlFormulas2 _
        , Lookat:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    ActiveCell.Offset(1, 0).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
'Set a variable so we can return to this regardless of filename
    Dim blndSht As String
    blndSht = ActiveWorkbook.Name

'Return to blend sheet workbook
    Windows(blndSht).Activate
    Application.CutCopyMode = False

End Sub


Sub blndShtGen2wk()
'///Assigned to the 2wk column on Blend These//////////////////////////////////////////////////////////////////////////////
'///Opens the blend sheet workbook for the blend on the row you click. Inputs the 2wk qty into said blend sheet.///////////

'String for current workbook name
    Dim src As String
    src = ActiveWorkbook.Name

'Copy the value of the shortage qty cell
    Selection.Copy
    
'Get the filepath + name of the appropriate blend sheet
    Dim blndShtPath As String
    blndShtPath = ActiveCell.Offset(0, -2).Value
    Workbooks.Open Filename:= _
        blndShtPath
    
'Paste the shortage qty value into the theory gallons cell
    Cells.Find(What:="Theory Gallons", After:=ActiveCell, LookIn:=xlFormulas2 _
        , Lookat:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    ActiveCell.Offset(1, 0).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
'Set a variable so we can return to this regardless of filename
    Dim blndSht As String
    blndSht = ActiveWorkbook.Name

'Return to blend sheet workbook
    Windows(blndSht).Activate
    Application.CutCopyMode = False

End Sub


Sub blndShtGen1wk()
'///Assigned to the 1wk column on Blend These//////////////////////////////////////////////////////////////////////////////
'///Opens the blend sheet workbook for the blend on the row you click. Inputs the 1wk qty into said blend sheet.///////////

'String for current workbook name
    Dim src As String
    src = ActiveWorkbook.Name

'Copy the value of the shortage qty cell
    Selection.Copy
    
'Get the filepath + name of the appropriate blend sheet
    Dim blndShtPath As String
    blndShtPath = ActiveCell.Offset(0, -1).Value
    Workbooks.Open Filename:= _
        blndShtPath
    
'Paste the shortage qty value into the theory gallons cell
    Cells.Find(What:="Theory Gallons", After:=ActiveCell, LookIn:=xlFormulas2 _
        , Lookat:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    ActiveCell.Offset(1, 0).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
'Set a variable so we can return to this regardless of filename
    Dim blndSht As String
    blndSht = ActiveWorkbook.Name

'Return to blend sheet workbook
    Windows(blndSht).Activate
    Application.CutCopyMode = False

End Sub


Sub LotNumGen()
'///Linked to T:T on BlendThese////////////////////////////////////////////////////////////////////////////////////////////
'///Opens Lot Number Generator.xlsx and fills in the blend desc////////////////////////////////////////////////////////////

'Set the way back to the current workbook
    Dim src As String
    src = ActiveWorkbook.Name

Call macrosOff

'Copy blend qty
    ActiveCell.Offset(0, -2).Select
    Selection.Copy
    
Call macrosOn

'Get filepath of Blenging Lont Ntumbner Grenerartror dot xlsx
    Dim path As String
    path = "C:\OD\Kinpak, Inc\Blending - Documents\01 Spreadsheet Tools\Blending Lot Number Generator\LotNumGenerator-Prod\Blending Lot Number Generator"
    Workbooks.Open path
    Sheets(1).Activate

'Paste the blend qty
    Range("D2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False

'Copy blend desc
    Windows(src).Activate
    ActiveCell.Offset(0, -13).Select
    Selection.Copy
    
'Paste blend desc
    Workbooks.Open path
    Range("B2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False

End Sub


Sub HotRoomSheet()
'///Linked to the fire icon on the IssueSheetTable/////////////////////////////////////////////////////////////////////////
'///Filters the IssueSheetTable Sheet to only include blends that need to be heated. Then prints the resulting sheet///////

'Off with the macros
    Call macrosOff

'Filter the sheet
    ActiveSheet.ListObjects("issueSheetTable").Range.AutoFilter Field:=10, _
        Criteria1:="HotRoom"
        
'Add row, for exposition
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "These blends must be heated TODAY so they will be ready for tomorrow."
    With Selection.Font
        .Name = "Calibri"
        .Size = 22
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
        
'Print that thang
    Call printIssueSchedule

'Delete row
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp

'Unfilter the sheet
    ActiveSheet.ListObjects("issueSheetTable").Range.AutoFilter Field:=10
    
'Commence to normalcy with the macros
    Call macrosOn

End Sub
