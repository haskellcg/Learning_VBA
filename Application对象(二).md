 1 Option Explicit
 2 
 3 'Application.OnTime
 4 'Application.OnKey
 5 'Application.WorksheetFunction
 6 'Application.Goto
 7 'Application.Union
 8 
 9 '绑定按键
10 Sub BindingKey()
11     Application.OnKey "%.", "NextPage"
12     Application.OnKey "%,", "PrevPage"
13 End Sub
14 
15 '向下
16 Sub NextPage()
17     ActiveWindow.LargeScroll down:=1
18 End Sub
19 
20 '向上
21 Sub PrevPage()
22     ActiveWindow.LargeScroll up:=1
23 End Sub
24 
25 '解除按键绑定
26 Sub UnbindingKey()
27     Application.OnKey "%."
28     Application.OnKey "%,"
29 End Sub
30 
31 '内置函数
32 'Application.WorksheetFunction.VLookup
33 'Application.WorksheetFunction.CountIf