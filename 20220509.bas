Attribute VB_Name = "Module6"
'多工作表資料視覺化
Sub autoVIsualizationSheets()
Dim shtIdx As Integer
For shtIdx = 1 To Sheets.Count
    Sheets(shtIdx).Activate
    Dim dtRange As Range
    Set dtRange = ActiveSheet.UsedRange
    Sheets(shtIdx).Select
    Range("A1").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.SetSourceData Source:=dtRange
    ActiveChart.ChartType = xlArea
Next
End Sub


Sub 錄製繪圖()
Dim ChartType As Integer
Select Case InputBox("要哪種圖形")

Case "圓餅圖"
ChartType = 5

Case "橫條圖"
ChartType = 57

Case "直條圖"
ChartType = 51

Case "折線圖"
ChartType = 4
End Select
Dim shtIdx As Integer
For shtIdx = 1 To Sheets.Count
    Sheets(shtIdx).Activate
    Dim dtRange As Range
    Set dtRange = ActiveSheet.UsedRange
    Sheets(shtIdx).Select
    Range("A1").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    ActiveSheet.Shapes.AddChart2(201, ChartType).Select
    ActiveChart.SetSourceData Source:=dtRange
    
Next

End Sub
Sub 巨集1()
Attribute 巨集1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 巨集1 巨集
'

'
    Range("A1:B14").Select
    ActiveSheet.Shapes.AddChart2(227, 4).Select
    ActiveChart.SetSourceData Source:=Range("分公司3!$A$1:$B$14")
End Sub
