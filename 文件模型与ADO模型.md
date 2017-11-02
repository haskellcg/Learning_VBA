 1 Option Explicit
 2 
 3 '文件操作函数记录
 4 'Dir
 5 'Name
 6 'Kill
 7 'RmDir
 8 'FileCopy
 9 'ChDir
10 'ChDrive
11 'MkDir
12 'RmDir
13 'CurDir
14 
15 'FreeFile
16 'Open
17 'Close
18 'Print
19 'Write
20 'Input
21 'EOF
22 'Line Input
23 
24 '文件对象模型
25 'File System Object, Scrrun.dll
26 'FileSystemObject
27 'Drive
28 'Folder
29 'File
30 'TextStream
31 
32 'Dim fsc as New Scripting.FileSystemObject
33 
34 'Dim fsc as Scriptinh.FileSystemObject
35 'Set fsc = New Scripting.FileSystemObject
36 
37 '获取文件信息
38 'FileLen
39 'GetAttr
40 'FileDateTime
41 'File.Attribute
42 
43 '文件管理
44 'FileExist
45 
46 'iFNumber = FreeFile
47 'Open sFName For Input As #iFNumber
48 'Line Input #iFNumber, strContent
49 
50 'Dim ostream As TextStream
51 'Set ostream = fso.OpenTextFile(Filename:=sFName, IOMode:=ForReading)
52 'strConent = ostream.ReadLine
53 
54 'ADO访问数据库
55 'Microsoft ActiveX Data Object Library
56 'Connection
57 'Connection.ConnectionString
58 'Connection.Open
59 'Connection.Execute
60 'RecordSet
61 'Set RecordSet = connection.open
62 'RecordSet.Open
63 'RecordSet.AddNew
64 'RecordSet.Delete
65 'RecordSet.Update
66 'RecordSet.Filed(index)
67 'RecordSet.MoveFirst
68 'Command
69 'Error
70 'Parameter
71 'Filed
72 
73 '步骤
74 'strCon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='"
75 'strCon = strCon & ActiveWorkBook.Path & "\人事管理.accdb"
76 'cnn.ConnectionString= strCon
77 'cnn.Open
78 'Set rst = cnn.Execute(strSql)
79 
80 'ADO访问Excel
81 'strCon = "Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties = Excel 8.0;" _
82 '           Data Source = 工作簿名称
83 
84 'ADO访问access
85 'strCon = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & _
86 '           ThisWorkbook.Path & "\Northwind.mdb"
87 
88 'Internet
89 'HyperLinks
90 'HyperLinks.Add
91 'HyperLinks.AddToFavorite
92 'HyperLinks.FollowHyperLink
93 'QueryTable
94 'QueryTable.Add
95 'QueryTable.Refresh
96 'PublicObject
97 'PublicObject.Add
98 'PublicObject.Publish