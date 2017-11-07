```VBA
Option Explicit

'Application.OnTime
'Application.OnKey
'Application.WorksheetFunction
'Application.Goto
'Application.Union

'绑定按键
Sub BindingKey()
    Application.OnKey "%.", "NextPage"
    Application.OnKey "%,", "PrevPage"
End Sub

'向下
Sub NextPage()
    ActiveWindow.LargeScroll down:=1
End Sub

'向上
Sub PrevPage()
    ActiveWindow.LargeScroll up:=1
End Sub

'解除按键绑定
Sub UnbindingKey()
    Application.OnKey "%."
    Application.OnKey "%,"
End Sub

'内置函数
'Application.WorksheetFunction.VLookup
'Application.WorksheetFunction.CountIf
```
