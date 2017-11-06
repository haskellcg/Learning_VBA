```VBA
Option Explicit
Option Base 1

Sub FillSheet()
    Dim i As Long
    Dim j As Long
    Dim col As Long
    Dim row As Long
    Dim arr() As Long
    row = Application.InputBox(prompt:="input row:", Type:=2)
    col = Application.InputBox(prompt:="input column:", Type:=2)
    ReDim arr(row, col)
    For i = 1 To row
        For j = 1 To col
            arr(i, j) = (i * j) Mod 20
        Next
    Next
    
    '需要选中
    Worksheets("Sheet3").Activate
    Dim rng As Variant
    Set rng = Sheets(3).Range(Cells(1, 1), Cells(row, col))
    'Set rng = Sheets(2).Range("A1:Z26")
    rng.Value = arr
End Sub
```
