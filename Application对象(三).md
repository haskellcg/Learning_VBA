这一部分学习事件的编写，分为两部分；

1、添加类，并编写如下代码；

　  这一部分用于编写事件功能，所有事件均在这里实现，代码如下：
   
```VBA
Option Explicit

Public WithEvents App As Application

Private Sub App_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    MsgBox "Sheet_Before_Double_Click", vbOKCancel
End Sub

Private Sub App_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    MsgBox "Sheet_Selection_Change", vbYesNoCancel
End Sub
```
	
2、添加模块，并编写如下代码；

　　这一部分用于链接虚参对象与全局对象的链接，代码如下：
  
```VBA
Option Explicit

Dim X As New EventClass

Sub ConnectApp()
    Set X.App = Application
End Sub

Sub DisConnectApp()
    Set X.App = Nothing
End Sub
```
