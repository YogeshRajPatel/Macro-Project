# Macro-Project
Streamlined the process of sharing monthly scorecards with specific individuals within the team.

Sub Spliter()

    Application.ScreenUpdating = False
    
    'AutoFit columns in "Raw" worksheet
    Worksheets("Raw").Columns.AutoFit
    
    'Open the destination workbook and copy data
    Dim wb As Workbook
    Set wb = Workbooks.Open("'INPUT FILE LOCATION'")
    wb.Worksheets("'INPUT SHEETNAME'").Activate
    Cells.Select
    Selection.Copy
    Workbooks("Scorecard Macro").Worksheets("Raw").Activate
    ActiveSheet.Paste
    
    'Define worksheet variables
    Dim Raw_sh As Worksheet
    Set Raw_sh = ThisWorkbook.Sheets("Raw")
    Dim spliter_sh As Worksheet
    Set spliter_sh = ThisWorkbook.Sheets("Spliter")
    Dim Filter_sh As Worksheet
    Set Filter_sh = ThisWorkbook.Sheets("Filter")
    
    'Clear existing data and apply filters
    Filter_sh.Range("A:A").Clear
    Raw_sh.AutoFilterMode = False
    Raw_sh.Range("E:E").Copy Filter_sh.Range("AL")
    Filter_sh.Range("A:A").RemoveDuplicates 1, xlYes
    Filter_sh.Range("A2").Clear
    Filter_sh.Range("B1").Value = Application.CountA(Filter_sh.Range("A:A"))
    
    'Loop through filtered data
    Dim i As Integer
    For i = 3 To Filter_sh.Range("B1").Value
        Raw_sh.UsedRange.AutoFilter Field:=2, Criteria1:="=" & Filter_sh.Range("A" & i).Value, Operator:=xlOr, Criteria2:="skill"
        Set nwb = Workbooks.Add
        Set nsh = nwb.Sheets(1)
        Raw_sh.UsedRange.SpecialCells(xlCellTypeVisible).Copy nsh.Range("AL")
        nsh.UsedRange.EntireColumn.ColumnWidth = 15
        nsh.Range("BL").Clear
        nwb.SaveAs spliter_sh.Range("C8").Value & "\" & Filter_sh.Range("A" & i).Value & ".xlsx"
        nwb.Close False
        Raw_sh.AutoFilterMode = False
    Next i
    
    'Message box
    MsgBox "All files successfully split. Click on Send Email button to send files to intended recipients"
    
    Application.ScreenUpdating = True
End Sub






Sub SendEmail()
    Application.ScreenUpdating = False
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets("Spliter")
    Dim i As Integer
    Dim OA As Object
    Dim msg As Object
    Set OA = CreateObject("Outlook.Application")
    Dim lastRow As Integer
    lastRow = sh.Range("C" & Application.Rows.Count).End(xlUp).Row
    For i = 15 To lastRow
        Set msg = OA.CreateItem(0)
        If sh.Range("A" & i).Value <> "" Then msg.SentOnBehalfOfName = sh.Range("A" & i).Value
        msg.To = sh.Range("B" & i).Value
        msg.Subject = sh.Range("C" & i).Value
        msg.Body = sh.Range("D" & i).Value
        msg.Attachments.Add sh.Range("E" & i).Value
        msg.Send
    Next i
    MsgBox "All emails sent."
    Application.ScreenUpdating = True
End Sub
