 1 Option Explicit
 2 
 3 Sub ExcelClass()
 4     
 5     'A1
 6     Dim oRng As Range
 7     Set oRng = Worksheets("Sheet1").Range("A1")
 8     oRng.Value = "地址"
 9     oRng.Font.Name = "楷体"
10     oRng.Font.Bold = True
11     
12     'array
13     Const FirstDimesion = 10
14     Const SecondDimesion = 10
15     Dim arrayRng(FirstDimesion, SecondDimesion) As Range
16     Dim i As Integer
17     Dim j As Integer
18     For i = 1 To FirstDimesion
19         For j = 1 To SecondDimesion
20             Set arrayRng(i, j) = Worksheets("Sheet1").Cells(i, j)
21         Next
22     Next
23     
24     Dim arrayItem As Range
25     Randomize
26     For i = 1 To FirstDimesion
27         For j = 1 To SecondDimesion
28             Set arrayItem = arrayRng(i, j)
29             arrayItem.Clear
30             arrayItem.AddComment ("this is comment")
31             arrayItem.Value = Rnd
32             arrayItem.Font.Name = "楷体"
33             arrayItem.Font.Color = RGB(128, 0, 128)
34             arrayItem.Font.Italic = True
35         Next
36     Next
37     
38     'Axes
39     'Charts
40     'Sheets    
41     'WorkBook
42     Dim sheetNew As Worksheet
43     Dim cellA1 As Range
44     Set sheetNew = Application.Workbooks(1).Worksheets.Add
45     Set cellA1 = sheetNew.Range("A1")
46     cellA1.Value = "中秋吃月饼"
47     cellA1.Font.Name = "nice"
48     cellA1.Font.Background = xlBackgroundTransparent
49 End Sub