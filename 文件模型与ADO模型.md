```VBA
Option Explicit

'文件操作函数记录
'Dir
'Name
'Kill
'RmDir
'FileCopy
'ChDir
'ChDrive
'MkDir
'RmDir
'CurDir

'FreeFile
'Open
'Close
'Print
'Write
'Input
'EOF
'Line Input

'文件对象模型
'File System Object, Scrrun.dll
'FileSystemObject
'Drive
'Folder
'File
'TextStream

'Dim fsc as New Scripting.FileSystemObject

'Dim fsc as Scriptinh.FileSystemObject
'Set fsc = New Scripting.FileSystemObject

'获取文件信息
'FileLen
'GetAttr
'FileDateTime
'File.Attribute

'文件管理
'FileExist

'iFNumber = FreeFile
'Open sFName For Input As #iFNumber
'Line Input #iFNumber, strContent

'Dim ostream As TextStream
'Set ostream = fso.OpenTextFile(Filename:=sFName, IOMode:=ForReading)
'strConent = ostream.ReadLine

'ADO访问数据库
'Microsoft ActiveX Data Object Library
'Connection
'Connection.ConnectionString
'Connection.Open
'Connection.Execute
'RecordSet
'Set RecordSet = connection.open
'RecordSet.Open
'RecordSet.AddNew
'RecordSet.Delete
'RecordSet.Update
'RecordSet.Filed(index)
'RecordSet.MoveFirst
'Command
'Error
'Parameter
'Filed

'步骤
'strCon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='"
'strCon = strCon & ActiveWorkBook.Path & "\人事管理.accdb"
'cnn.ConnectionString= strCon
'cnn.Open
'Set rst = cnn.Execute(strSql)

'ADO访问Excel
'strCon = "Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties = Excel 8.0;" _
'           Data Source = 工作簿名称

'ADO访问access
'strCon = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & _
'           ThisWorkbook.Path & "\Northwind.mdb"

'Internet
'HyperLinks
'HyperLinks.Add
'HyperLinks.AddToFavorite
'HyperLinks.FollowHyperLink
'QueryTable
'QueryTable.Add
'QueryTable.Refresh
'PublicObject
'PublicObject.Add
'PublicObject.Publish
```
