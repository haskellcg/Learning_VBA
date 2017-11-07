```VBA
Option Explicit

'有返回值时需要加括号
'没有返回值则不需要加括号

'显示单个字符
Sub ShowCharWithDuring(strData As String, during As Single)
    Application.SendKeys strData, True
    Dim startTime As Single
    startTime = Timer
    Do While Timer < (startTime + during)
        DoEvents
    Loop
End Sub

'显示字符串
Sub ShowStringWithDuring(strData As String, during As Single)
    Dim strLen As Integer
    Dim i As Integer
    Dim strChar As String
    strLen = Len(strData)
    For i = 1 To strLen
        strChar = Mid(strData, i, 1)
        ShowCharWithDuring strChar, during
    Next
End Sub

Sub ApplicationClass()
    Dim notepadHandle As Double
    notepadHandle = Shell("E:\Program Files (x86)\Notepad++\notepad++", 1)
    AppActivate notepadHandle
    Application.SendKeys "~", True
    Application.SendKeys "Keybord Input Demo.", True
    Application.SendKeys "~", True
    Application.SendKeys "Excel 2010 VBA.", True
    Application.SendKeys "~", True
    
    '逐字显示文章
    Dim strEssay As String
    strEssay = "* ############################ *~" & _
               "* This is my comment.          *~" & _
               "* We use notepad to show info. *~" & _
               "* ############################ *~"
    ShowStringWithDuring strEssay, 0.2
End Sub
```
