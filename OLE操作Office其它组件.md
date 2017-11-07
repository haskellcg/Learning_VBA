```VBA
Option Explicit

'Object Link Emdeded

'前期绑定
'Dim myWord As Word.Application
'Set myWord = New Word.Application

'后期绑定
'Dim myWord As Object
'Set myWord = CreateObject("Word.Application")
'Set myWord = CreateObject("PowerPoint.Application")

'Word对象模型
'Application
'Bookmark
'Document
'Paragraph
'Range
'Selection


Sub OpenWordApp()
    Dim strFileName As String
    Dim strFilter As String
    Dim strTitle As String
    Dim appDoc As Object
    
    strFilter = "Word文档(*.doc;*.docx;*.docm),*.doc;*.docx;*.docm"
    strTitle = "Open Word Document"
    
    strFileName = Application.GetOpenFilename _
                    (filefilter:=strFilter, _
                    Title:=strTitle)
    If strFileName = "False" Then
        Exit Sub
    End If
    
    Set appDoc = CreateObject("Word.Application")
    '打开文档
    appDoc.Documents.Open strFileName
    '设置可见
    appDoc.Visible = True
    
    Dim i As Integer
    Dim pg As Variant
    Dim strContent As String
    i = 1
    With appDoc.ActiveDocument
        For Each pg In .Paragraphs
            strContent = pg.Range.Text
            
            i = i + 1
            ActiveSheet.Cells(i, 1) = strContent
        Next
    End With
    
    appDoc.Visible = False
    appDoc.Quit
    Set appDoc = Nothing
End Sub

'PowerPoint对象模型
'Application
'Presentation
'Slide
'Slides

'Outlook对象模型
'Application
'Explorer
'Inspector
'Folder
'MailItem
```
