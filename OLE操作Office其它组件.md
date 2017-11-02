1 Option Explicit
 2 
 3 'Object Link Emdeded
 4 
 5 '前期绑定
 6 'Dim myWord As Word.Application
 7 'Set myWord = New Word.Application
 8 
 9 '后期绑定
10 'Dim myWord As Object
11 'Set myWord = CreateObject("Word.Application")
12 'Set myWord = CreateObject("PowerPoint.Application")
13 
14 'Word对象模型
15 'Application
16 'Bookmark
17 'Document
18 'Paragraph
19 'Range
20 'Selection
21 
22 
23 Sub OpenWordApp()
24     Dim strFileName As String
25     Dim strFilter As String
26     Dim strTitle As String
27     Dim appDoc As Object
28     
29     strFilter = "Word文档(*.doc;*.docx;*.docm),*.doc;*.docx;*.docm"
30     strTitle = "Open Word Document"
31     
32     strFileName = Application.GetOpenFilename _
33                     (filefilter:=strFilter, _
34                     Title:=strTitle)
35     If strFileName = "False" Then
36         Exit Sub
37     End If
38     
39     Set appDoc = CreateObject("Word.Application")
40     '打开文档
41     appDoc.Documents.Open strFileName
42     '设置可见
43     appDoc.Visible = True
44     
45     Dim i As Integer
46     Dim pg As Variant
47     Dim strContent As String
48     i = 1
49     With appDoc.ActiveDocument
50         For Each pg In .Paragraphs
51             strContent = pg.Range.Text
52             
53             i = i + 1
54             ActiveSheet.Cells(i, 1) = strContent
55         Next
56     End With
57     
58     appDoc.Visible = False
59     appDoc.Quit
60     Set appDoc = Nothing
61 End Sub
62 
63 'PowerPoint对象模型
64 'Application
65 'Presentation
66 'Slide
67 'Slides
68 
69 'Outlook对象模型
70 'Application
71 'Explorer
72 'Inspector
73 'Folder
74 'MailItem