 1 Option Explicit
 2 Option Base 1
 3 
 4 Sub FillSheet()
 5     Dim i As Long
 6     Dim j As Long
 7     Dim col As Long
 8     Dim row As Long
 9     Dim arr() As Long
10     row = Application.InputBox(prompt:="input row:", Type:=2)
11     col = Application.InputBox(prompt:="input column:", Type:=2)
12     ReDim arr(row, col)
13     For i = 1 To row
14         For j = 1 To col
15             arr(i, j) = (i * j) Mod 20
16         Next
17     Next
18     
19     '需要选中
20     Worksheets("Sheet3").Activate
21     Dim rng As Variant
22     Set rng = Sheets(3).Range(Cells(1, 1), Cells(row, col))
23     'Set rng = Sheets(2).Range("A1:Z26")
24     rng.Value = arr
25 End Sub