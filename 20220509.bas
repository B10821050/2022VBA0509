Attribute VB_Name = "Module6"
'�h�u�@���Ƶ�ı��
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


Sub ���sø��()
Dim ChartType As Integer
Select Case InputBox("�n���عϧ�")

Case "����"
ChartType = 5

Case "�����"
ChartType = 57

Case "������"
ChartType = 51

Case "��u��"
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
Sub ����1()
Attribute ����1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ����1 ����
'

'
    Range("A1:B14").Select
    ActiveSheet.Shapes.AddChart2(227, 4).Select
    ActiveChart.SetSourceData Source:=Range("�����q3!$A$1:$B$14")
End Sub
