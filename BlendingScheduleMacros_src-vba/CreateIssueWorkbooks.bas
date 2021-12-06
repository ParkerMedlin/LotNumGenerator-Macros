Sub createWorkbooks()
'///Linked to H1 on issueSheetTable sheet///////////////////////////////////////////////////////////////////////////////////////////////
'///Loops through the list and triggers IssueSheetGen at each row///////////////////////////////////////////////////////////////////////

'msgbox to make sure u committed bb <3
    Dim uInput2 As Integer
    uInput2 = MsgBox("Proceed with all Issue Sheet Functions?", vbQuestion + vbYesNo)
    
    If uInput2 = vbYes Then
    
'turn the macros off
        Call macrosOff
        
'Set the way back to the current workbook
        Dim src As String
        src = ActiveWorkbook.Name

'Get the date and put it into NOT the Blending Issue Sheet
        Dim issueDate As Date
        issueDate = InputBox("Enter the date as MM/DD/YYYY")
        ChDir "G:\Blending\02 Blending Issue Sheet"
        Workbooks.Open Filename:= _
        "G:\Blending\02 Blending Issue Sheet\NOT the Blending Issue Sheet.xlsb"
        Range("D4").Value = issueDate

        Windows(src).Activate
        
'Remove duplicates from the list
        Range("issueSheetTable[uniqchek]").Select
        ActiveSheet.Range("issueSheetTable[#All]").RemoveDuplicates Columns:=15, _
            Header:=xlYes
    
'Declare variables
        Dim numberOfRows As Integer
        Dim i As Integer
   
'Set the number of total iterations to the number of rows in the issue sheet table
        ActiveSheet.ListObjects("issueSheetTable").DataBodyRange.SpecialCells(xlCellTypeVisible).Select
        numberOfRows = Selection.Rows.Count
        Range("H2").Select
    
'The For loop
        For i = 1 To numberOfRows
        Call IssueSheetGen
        Selection.Offset(1, -1).Select
        Next i
    
'Print and save issue sheets
        Call saveNewIssueSheets
    
'Close the template
        Workbooks("NOT the Blending Issue Sheet.xlsb").Close SaveChanges:=False
    
'Refresh and restore order
        Call refreshIssueSheet

'turn the macros on
        Call macrosOn
        Range("I2:I40").Select
        Selection.ClearContents
        
    
'Print the issue schedule
        Dim uInput3 As Integer
        uInput3 = MsgBox("Print Issue Schedules x6?", vbQuestion + vbYesNo)
    
        If uInput3 = vbYes Then
            
            Call printIssueSchedule
    
        Else
        MsgBox "Print cancelled."
        End If
   
    Else
        MsgBox "Looper cancelled."
    End If

End Sub

Sub IssueSheetGen()
'///Linked to H:H on issueSheetTable sheet//////////////////////////////////////////////////////////////////////////////////////////////
'///Generates an issue sheet for the blend PN of the row that is clicked////////////////////////////////////////////////////////////////

'Set the way back to the current workbook
    Dim src As String
    src = ActiveWorkbook.Name

'Open the blend issue sheet
    ChDir "G:\Blending\02 Blending Issue Sheet"
    Workbooks.Open Filename:= _
        "G:\Blending\02 Blending Issue Sheet\NOT the Blending Issue Sheet.xlsb"

'Go back to the original workbook (Blending Schedule) and copy the production line name
    Windows(src).Activate
        ActiveCell.Offset(0, -2).Select
        Selection.Copy
        
'Go to blend issue sheet workbook and paste the production line name
    Windows("NOT the Blending Issue Sheet.xlsb").Activate
        Range("$D$6").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

'Go to the original workbook (Blending Schedule) and copy the blend desc
    Windows(src).Activate
        ActiveCell.Offset(0, -3).Select
        Application.CutCopyMode = False
        Selection.Copy
        
'Go to blend issue sheet workbook and paste the blend desc
    Windows("NOT the Blending Issue Sheet.xlsb").Activate
        Range("$D$8").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
            
'Copy the Blend Issue Sheet worksheet to the appropriate Line worksheet
    Call workSheetCopy
    
'mark x for done
    Windows(src).Activate
        ActiveCell.Offset(0, 6).Select
        Selection.Value = "x"
 
End Sub


Sub workSheetCopy()
'///Copies the issue sheets to the appropriate issue sheet template depending on which line it is///////////////////////////////////////

'SET THE PATH BACK, MY LAD
    Dim src As String
    src = "NOT the Blending Issue Sheet.xlsb"
    Dim SheetCount As Integer
    Dim BlendName As String
    BlendName = Range("B8").Value
    Dim dest As String

'Declare the path for the workbook I'm copying all the Inline issue sheets to
    Dim InlinePath As String
    InlinePath = "C:\OD\Kinpak, Inc\Blending - Documents\01 Spreadsheet Tools\Blending Issue Sheet\PrintToSheet\01 Inline Issue Sheets.xlsb"

'Declare the path for the workbook I'm copying all the PDline issue sheets to
    Dim PDlinePath As String
    PDlinePath = "C:\OD\Kinpak, Inc\Blending - Documents\01 Spreadsheet Tools\Blending Issue Sheet\PrintToSheet\02 PDline Issue Sheets.xlsb"

'Declare the path for the workbook I'm copying all the JBline issue sheets to
    Dim JBlinePath As String
    JBlinePath = "C:\OD\Kinpak, Inc\Blending - Documents\01 Spreadsheet Tools\Blending Issue Sheet\PrintToSheet\03 JBline Issue Sheets.xlsb"

'Select the appropriate workbook and copy the issue sheet to it
    If Range("D6").Value = "Inline" Then
        Workbooks.Open (InlinePath)
        SheetCount = Application.Sheets.Count
        Windows(src).Activate
        Sheets("Blending Issue Sheet").Select
        Sheets("Blending Issue Sheet").Copy Before:=Workbooks( _
        "01 Inline Issue Sheets.xlsb").Sheets(SheetCount)
        Windows("01 Inline Issue Sheets.xlsb").Activate
        ActiveSheet.Name = BlendName
        Call copyBatchNumberValues

    ElseIf Range("D6").Value = "PDLine" Then
        Workbooks.Open (PDlinePath)
        SheetCount = Application.Sheets.Count
        Windows(src).Activate
        Sheets("Blending Issue Sheet").Select
        Sheets("Blending Issue Sheet").Copy Before:=Workbooks( _
        "02 PDline Issue Sheets.xlsb").Sheets(SheetCount)
        Windows("02 PDline Issue Sheets.xlsb").Activate
582        ActiveSheet.Name = BlendName
        Call copyBatchNumberValues

    ElseIf Range("D6").Value = "JBLine" Then
        Workbooks.Open (JBlinePath)
        SheetCount = Application.Sheets.Count
        Windows(src).Activate
        Sheets("Blending Issue Sheet").Select
        Sheets("Blending Issue Sheet").Copy Before:=Workbooks( _
        "03 JBline Issue Sheets.xlsb").Sheets(SheetCount)
        Windows("03 JBline Issue Sheets.xlsb").Activate
        ActiveSheet.Name = BlendName
        Call copyBatchNumberValues
    
    End If


End Sub


Sub copyBatchNumberValues()
'///Copies the batch numbers and pastes them as values in the appropriate slots in the new thing////////////////////////////////////////

    Dim dest As String
    lineIssueShtWorkbook = ActiveWorkbook.Name
    
    Windows("NOT the Blending Issue Sheet.xlsb").Activate
    Range("A4:F27").Select
    Selection.Copy
    Windows(lineIssueShtWorkbook).Activate
    Range("A4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
      :=False, Transpose:=False
    Windows("NOT the Blending Issue Sheet.xlsb").Activate
    
    Application.CutCopyMode = False
    
End Sub


Sub saveNewIssueSheets()
'///Save a copy of each issue sheet and rename them appropriately///////////////////////////////////////////////////////////////////////

'declare variables for each filename
    Dim INpath As String
    Dim PDpath As String
    Dim JBpath As String
    Dim foldrpath As String
    Dim dateStr As String
    
    dateStr = Format(Date + 1, "mm-dd-yyyy")

'Create folder in OneDrive\Blending\DatedIssueSheets
    foldrpath = "C:\OD\Kinpak, Inc\Blending - Documents\01 Spreadsheet Tools\Blending Issue Sheet\IssueSheetArchive\"

    foldrpath = foldrpath & dateStr & " Issue Sheets"
    MsgBox foldrpath
    MkDir foldrpath
    
'Open and save copy of each set of sheets to a new workbook, then clear the template:
    INpath = foldrpath & "\" & dateStr & " Inline Issue Sheets.xlsb"
    Windows("01 Inline Issue Sheets.xlsb").Activate
    ActiveWorkbook.SaveCopyAs INpath
    Dim uInput As Integer
    uInput = MsgBox("Print Issue Sheets?", vbQuestion + vbYesNo)
        If uInput = vbYes Then
            Call printIssueSheets
            Call printIssueSheets
        Else
        MsgBox "Print cancelled."
        End If
    Workbooks("01 Inline Issue Sheets.xlsb").Close SaveChanges:=False
    
    PDpath = foldrpath & "\" & dateStr & " PDline Issue Sheets.xlsb"
    Windows("02 PDline Issue Sheets.xlsb").Activate
    ActiveWorkbook.SaveCopyAs PDpath
    uInput = MsgBox("Print Issue Sheets?", vbQuestion + vbYesNo)
        If uInput = vbYes Then
            Call printIssueSheets
            Call printIssueSheets
        Else
        MsgBox "Print cancelled."
        End If
    Workbooks("02 PDline Issue Sheets.xlsb").Close SaveChanges:=False

    JBpath = foldrpath & "\" & dateStr & " JBline Issue Sheets.xlsb"
    Windows("03 JBline Issue Sheets.xlsb").Activate
    ActiveWorkbook.SaveCopyAs JBpath
    uInput = MsgBox("Print Issue Sheets?", vbQuestion + vbYesNo)
        If uInput = vbYes Then
            Call printIssueSheets
            Call printIssueSheets
        Else
        MsgBox "Print cancelled."
        End If
    Workbooks("03 JBline Issue Sheets.xlsb").Close SaveChanges:=False
    
    
End Sub


