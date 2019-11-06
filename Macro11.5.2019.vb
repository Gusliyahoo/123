Public wbName As String
Public arr As Variant
Sub NewWorkbook()
    Application.ScreenUpdating = False
    Dim newwb As Workbook
    Set newwb = Workbooks.Add
    wbName = "Test " & Format(Now, "mm-dd-yyyy HH_mm_ss") & ".xls"
    ActiveWorkbook.SaveAs Filename:=wbName
    AddSheets
    AddContent
    Application.ScreenUpdating = True
End Sub
Sub AddSheets()
    Dim wb As Workbook: Set wb = Workbooks(wbName)
    Dim ws As Worksheet
    arr = Array("MGM", "ILO", "GUA")
    For Each Item In arr
        If Item = "MGM" Then
            Worksheets("Sheet1").Name = "MGM"
        Else
            Set ws = wb.Worksheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))
            With ws
                .Name = Item
            End With
        End If
    Next Item
End Sub
Sub AddContent()
    For Each Item In arr
        Windows("Book1.xlsm").Activate
        Sheets(Item).Activate
        Range("A1").Select
        Selection.CurrentRegion.Select
        Selection.Interior.Color = RGB(201, 255, 228)
        With Selection.Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        Selection.Copy
        Windows(wbName).Activate
        Sheets(Item).Activate
        ActiveSheet.Paste
    Next Item
End Sub
