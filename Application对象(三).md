这一部分学习事件的编写，分为两部分；

1、添加类，并编写如下代码；

　  这一部分用于编写事件功能，所有事件均在这里实现，代码如下：

	 1 Option Explicit
	 2 
	 3 Public WithEvents App As Application
	 4 
	 5 Private Sub App_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
	 6     MsgBox "Sheet_Before_Double_Click", vbOKCancel
	 7 End Sub
	 8 
	 9 Private Sub App_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
	10     MsgBox "Sheet_Selection_Change", vbYesNoCancel
	11 End Sub
	
2、添加模块，并编写如下代码；

　　这一部分用于链接虚参对象与全局对象的链接，代码如下：

	 1 Option Explicit
	 2 
	 3 Dim X As New EventClass
	 4 
	 5 Sub ConnectApp()
	 6     Set X.App = Application
	 7 End Sub
	 8 
	 9 Sub DisConnectApp()
	10     Set X.App = Nothing
	11 End Sub