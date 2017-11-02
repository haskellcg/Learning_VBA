 1 Option Explicit
 2 
 3 '有返回值时需要加括号
 4 '没有返回值则不需要加括号
 5 
 6 '显示单个字符
 7 Sub ShowCharWithDuring(strData As String, during As Single)
 8     Application.SendKeys strData, True
 9     Dim startTime As Single
10     startTime = Timer
11     Do While Timer < (startTime + during)
12         DoEvents
13     Loop
14 End Sub
15 
16 '显示字符串
17 Sub ShowStringWithDuring(strData As String, during As Single)
18     Dim strLen As Integer
19     Dim i As Integer
20     Dim strChar As String
21     strLen = Len(strData)
22     For i = 1 To strLen
23         strChar = Mid(strData, i, 1)
24         ShowCharWithDuring strChar, during
25     Next
26 End Sub
27 
28 Sub ApplicationClass()
29     Dim notepadHandle As Double
30     notepadHandle = Shell("E:\Program Files (x86)\Notepad++\notepad++", 1)
31     AppActivate notepadHandle
32     Application.SendKeys "~", True
33     Application.SendKeys "Keybord Input Demo.", True
34     Application.SendKeys "~", True
35     Application.SendKeys "Excel 2010 VBA.", True
36     Application.SendKeys "~", True
37     
38     '逐字显示文章
39     Dim strEssay As String
40     strEssay = "* ############################ *~" & _
41                "* This is my comment.          *~" & _
42                "* We use notepad to show info. *~" & _
43                "* ############################ *~"
44     ShowStringWithDuring strEssay, 0.2
45 End Sub