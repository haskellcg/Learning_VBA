```VBA
Option Explicit

Sub ExcelClass()
    
    'A1
    Dim oRng As Range
    Set oRng = Worksheets("Sheet1").Range("A1")
    oRng.Value = "地址"
    oRng.Font.Name = "楷体"
    oRng.Font.Bold = True
    
    'array
    Const FirstDimesion = 10
    Const SecondDimesion = 10
    Dim arrayRng(FirstDimesion, SecondDimesion) As Range
    Dim i As Integer
    Dim j As Integer
    For i = 1 To FirstDimesion
        For j = 1 To SecondDimesion
            Set arrayRng(i, j) = Worksheets("Sheet1").Cells(i, j)
        Next
    Next
    
    Dim arrayItem As Range
    Randomize
    For i = 1 To FirstDimesion
        For j = 1 To SecondDimesion
            Set arrayItem = arrayRng(i, j)
            arrayItem.Clear
            arrayItem.AddComment ("this is comment")
            arrayItem.Value = Rnd
            arrayItem.Font.Name = "楷体"
            arrayItem.Font.Color = RGB(128, 0, 128)
            arrayItem.Font.Italic = True
        Next
    Next
    
    'Axes
    'Charts
    'Sheets    
    'WorkBook
    Dim sheetNew As Worksheet
    Dim cellA1 As Range
    Set sheetNew = Application.Workbooks(1).Worksheets.Add
    Set cellA1 = sheetNew.Range("A1")
    cellA1.Value = "中秋吃月饼"
    cellA1.Font.Name = "nice"
    cellA1.Font.Background = xlBackgroundTransparent
End Sub
```
