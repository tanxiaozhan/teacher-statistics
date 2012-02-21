VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form mainfrm 
   Caption         =   "调查问卷统计系统"
   ClientHeight    =   4440
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7065
   LinkTopic       =   "Form1"
   ScaleHeight     =   4440
   ScaleWidth      =   7065
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command2 
      Caption         =   "退    出"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2070
      TabIndex        =   5
      Top             =   3480
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "选项"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2085
      TabIndex        =   3
      Top             =   2610
      Width           =   2775
   End
   Begin VB.FileListBox File1 
      Height          =   270
      Left            =   705
      TabIndex        =   2
      Top             =   3225
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSComctlLib.ProgressBar PBar 
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton tj 
      Caption         =   "开始统计"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2085
      TabIndex        =   0
      Top             =   1740
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   255
      TabIndex        =   6
      Top             =   120
      Width           =   6675
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3090
      TabIndex        =   4
      Top             =   630
      Width           =   1065
   End
End
Attribute VB_Name = "mainfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    optionfrm.Show vbModal, Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    File1.Path = docPath
End Sub

Private Sub Form_Load()
    dbName = App.Path
    If Right(dbName, 1) <> "\" Then dbName = dbName & "\"
    dbName = dbName & "data.mdb"
    Set conn = New ADODB.Connection
    docPath = App.Path & "\doc"
    File1.Path = docPath
    Label1.Caption = ""
    Label2.Caption = ""
End Sub

Private Sub tj_Click()
'On Error GoTo errmsg
    Dim errFileName(1000) As String
    Dim errFileNum As Integer
    Dim fileCount
    Dim wApp As Word.Application
    
    If Dir(dbName) = "" Then CreateDB
    
    If DirExists(App.Path & "\errFiles") = 0 Then
        MkDir App.Path & "\errFiles"
    End If
    
    
    
    DBConnect
    conn.Execute "delete from tj"
    
    Set wApp = New Word.Application
    
    wApp.Visible = False
    fileCount = File1.ListCount
    
    PBar.Min = 0
    PBar.Max = fileCount
    PBar.Value = PBar.Min
    errFileNum = 0
    For j = 0 To fileCount - 1
    Label1.Caption = File1.List(j)
        wApp.Documents.Open docPath & "\" & File1.List(j)
    'For j = 0 To 0
    'wApp.Documents.Open docPath & "\wj(4).doc"
        sql = "insert into tj(xx,xk,nl,t2e,t4,t5,t1a,t1b,t1c,t1d,t2a,t2b,t2c,t2d,t3a,t3b,t3c,t3d) values("
        fieldValue = ""
        If wApp.ActiveDocument.Shapes.count = 6 And wApp.ActiveDocument.InlineShapes.count = 13 Then
        
        '对于浮动式文本框控件,用shapes
        For i = 1 To wApp.ActiveDocument.Shapes.count
            'MsgBox wApp.ActiveDocument.Shapes(i).OLEFormat.Object
            fieldValue = fieldValue & "'" & wApp.ActiveDocument.Shapes(i).OLEFormat.Object & "',"
        Next
        fieldValue = Left(fieldValue, Len(fieldValue) - 1)
    
        '对于嵌入式文本框控件,用inlineshapes
        For i = 2 To wApp.ActiveDocument.InlineShapes.count
            fieldValue = fieldValue & "," & wApp.ActiveDocument.InlineShapes(i).OLEFormat.Object
        Next
        
    
        Label1.Caption = "问卷数量：" & j + 1 & "份"
        
        PBar.Value = PBar + 1
        Label2.Caption = Int(PBar.Value / PBar.Max * 100) & "%"
        sql = sql & fieldValue & ")"
        conn.Execute sql
    Else
        errFileNum = errFileNum + 1
        errFileName(errFileNum) = File1.List(j)
    
    End If
    
    wApp.ActiveDocument.Close
      
    Next
    
    
    Label2.Caption = "100%"
    
    Label1.Caption = "正在生成统计表..."
    
    tongji
        
    conn.Close
    Set conn = Nothing
    
errmsg:
    If wApp <> "" Then wApp.Quit
    Set wApp = Nothing
    
    Label1.Caption = "统计完成！保存到" & App.Path & "文件夹。"
    
    errFilemsg = ""
    For i = 1 To errFileNum
        FileCopy docPath & "\" & errFileName(i), App.Path & "\errFiles" & errFileName(i)
        errFilemsg = errFilemsg & errFileName(i) & Chr(13)
    Next
    errFilemsg = "下列文件未统计，已复制到" & App.Path & "\errFiles" & Chr(13) & Chr(13) & errFilemsg
    MsgBox errFilemsg, vbCritical, "错误文件列表"
    
End Sub

Private Sub CreateDB()
    '菜单“工程”-->"引用"-->"Microsoft   ActiveX   Data   Objects   2.8   Library"
    '                    -->  Microsoft   ADO   Ext.2.8   for   DDL   ado   Security
    Dim cat     As ADOX.Catalog
    Set cat = New ADOX.Catalog
    cat.Create ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbName & ";")
    MsgBox "数据库创建成功！"
    Dim tbl     As ADOX.Table
    Set tbl = New ADOX.Table
    tbl.ParentCatalog = cat
    tbl.Name = "tj"
    
    '增加一个自动增长的字段
    Dim col     As ADOX.Column
    Set col = New ADOX.Column
    col.ParentCatalog = cat
    col.Type = ADOX.DataTypeEnum.adInteger       '   //   必须先设置字段类型
    col.Name = "id"
    col.Properties("Jet OLEDB:Allow Zero Length").Value = False
    col.Properties("AutoIncrement").Value = True
    tbl.Columns.Append col, ADOX.DataTypeEnum.adInteger, 0
    
    '增加一个文本字段
    Dim col2     As ADOX.Column
    Set col2 = New ADOX.Column
    col2.ParentCatalog = cat
    col2.Name = "xx"   '学校名称
    col2.Properties("Jet OLEDB:Allow Zero Length").Value = True
    tbl.Columns.Append col2, ADOX.DataTypeEnum.adVarChar, 50
    
    '增加一个文本字段
    Dim col3     As ADOX.Column
    Set col3 = New ADOX.Column
    col3.ParentCatalog = cat
    col3.Name = "xk"   '学科
    col3.Properties("Jet OLEDB:Allow Zero Length").Value = True
    tbl.Columns.Append col3, ADOX.DataTypeEnum.adVarChar, 50
    
    '增加一个数值型字段
    Dim col4     As ADOX.Column
    Set col4 = New ADOX.Column
    col4.ParentCatalog = cat
    col4.Name = "nl"   '年龄
    tbl.Columns.Append col4, ADOX.DataTypeEnum.adVarChar, 5
    
    '增加一个数值型字段
    Dim col5     As ADOX.Column
    Set col5 = New ADOX.Column
    col5.ParentCatalog = cat
    col5.Type = ADOX.DataTypeEnum.adBoolean
    col5.Name = "t1a"   '1a
    tbl.Columns.Append col5, ADOX.DataTypeEnum.adBoolean
    
    '增加一个数值型字段
    Dim col6     As ADOX.Column
    Set col6 = New ADOX.Column
    col6.ParentCatalog = cat
    col6.Type = ADOX.DataTypeEnum.adBoolean
    col6.Name = "t1b"   '1b
    tbl.Columns.Append col6, ADOX.DataTypeEnum.adBoolean
    
    '增加一个数值型字段
    Dim col7     As ADOX.Column
    Set col7 = New ADOX.Column
    col7.ParentCatalog = cat
    col7.Type = ADOX.DataTypeEnum.adBoolean
    col7.Name = "t1c"   '1c
    tbl.Columns.Append col7, ADOX.DataTypeEnum.adBoolean
    
    '增加一个数值型字段
    Dim col8     As ADOX.Column
    Set col8 = New ADOX.Column
    col8.ParentCatalog = cat
    col8.Type = ADOX.DataTypeEnum.adBoolean
    col8.Name = "t1d"   '1d
    tbl.Columns.Append col8, ADOX.DataTypeEnum.adBoolean
    
    '增加一个数值型字段
    Dim col9     As ADOX.Column
    Set col9 = New ADOX.Column
    col9.ParentCatalog = cat
    col9.Type = ADOX.DataTypeEnum.adBoolean
    col9.Name = "t2a"   '2a
    tbl.Columns.Append col9, ADOX.DataTypeEnum.adBoolean
    
    '增加一个数值型字段
    Dim col10     As ADOX.Column
    Set col10 = New ADOX.Column
    col10.ParentCatalog = cat
    col10.Type = ADOX.DataTypeEnum.adBoolean
    col10.Name = "t2b"   '2b
    tbl.Columns.Append col10, ADOX.DataTypeEnum.adBoolean
    
    '增加一个数值型字段
    Dim col11     As ADOX.Column
    Set col11 = New ADOX.Column
    col11.ParentCatalog = cat
    col11.Type = ADOX.DataTypeEnum.adBoolean
    col11.Name = "t2c"   '2c
    tbl.Columns.Append col11, ADOX.DataTypeEnum.adBoolean
    
    '增加一个数值型字段
    Dim col12     As ADOX.Column
    Set col12 = New ADOX.Column
    col12.ParentCatalog = cat
    col12.Type = ADOX.DataTypeEnum.adBoolean
    col12.Name = "t2d"   '2d
    tbl.Columns.Append col12, ADOX.DataTypeEnum.adBoolean
    
    '增加一个文本字段
    Dim col13     As ADOX.Column
    Set col13 = New ADOX.Column
    col13.ParentCatalog = cat
    col13.Name = "t2e"   '2e
    col13.Properties("Jet OLEDB:Allow Zero Length").Value = True
    tbl.Columns.Append col13, ADOX.DataTypeEnum.adVarChar, 255
    
    '增加一个数值型字段
    Dim col14     As ADOX.Column
    Set col14 = New ADOX.Column
    col14.ParentCatalog = cat
    col14.Type = ADOX.DataTypeEnum.adBoolean
    col14.Name = "t3a"   '3a
    tbl.Columns.Append col14, ADOX.DataTypeEnum.adBoolean
    
    '增加一个数值型字段
    Dim col15     As ADOX.Column
    Set col15 = New ADOX.Column
    col15.ParentCatalog = cat
    col15.Type = ADOX.DataTypeEnum.adBoolean
    col15.Name = "t3b"   '3b
    tbl.Columns.Append col15, ADOX.DataTypeEnum.adBoolean
    
    '增加一个数值型字段
    Dim col16     As ADOX.Column
    Set col16 = New ADOX.Column
    col16.ParentCatalog = cat
    col16.Type = ADOX.DataTypeEnum.adBoolean
    col16.Name = "t3c"   '3c
    tbl.Columns.Append col16, ADOX.DataTypeEnum.adBoolean
    
    '增加一个数值型字段
    Dim col17     As ADOX.Column
    Set col17 = New ADOX.Column
    col17.ParentCatalog = cat
    col17.Type = ADOX.DataTypeEnum.adBoolean
    col17.Name = "t3d"   '3d
    tbl.Columns.Append col17, ADOX.DataTypeEnum.adBoolean

    '增加一个文本字段
    Dim col18     As ADOX.Column
    Set col18 = New ADOX.Column
    col18.ParentCatalog = cat
    col18.Name = "t4"   '4
    col18.Properties("Jet OLEDB:Allow Zero Length").Value = True
    tbl.Columns.Append col18, ADOX.DataTypeEnum.adVarChar, 255
    
    '增加一个文本字段
    Dim col19     As ADOX.Column
    Set col19 = New ADOX.Column
    col19.ParentCatalog = cat
    col19.Name = "t5"   '5
    col19.Properties("Jet OLEDB:Allow Zero Length").Value = True
    tbl.Columns.Append col19, ADOX.DataTypeEnum.adVarChar, 255
 
    
    '增加一个货币型字段
    'Dim col4     As ADOX.Column
    'Set col4 = New ADOX.Column
    'col4.ParentCatalog = cat
    'col4.Type = ADOX.DataTypeEnum.adCurrency
    'col4.Name = "xx"
    'tbl.Columns.Append col4, ADOX.DataTypeEnum.adCurrency
    
    '增加一个OLE字段
    'Dim col5     As ADOX.Column
    'Set col5 = New ADOX.Column
    'col5.ParentCatalog = cat
    'col5.Type = ADOX.DataTypeEnum.adLongVarBinary
    'col5.Name = "OLD_FLD"
    'tbl.Columns.Append col5, ADOX.DataTypeEnum.adLongVarBinary
    
    '增加一个数值型字段
    'Dim col3     As ADOX.Column
    'Set col3 = New ADOX.Column
    'col3.ParentCatalog = cat
    'col3.Type = ADOX.DataTypeEnum.adDouble
    'col3.Name = "ll"
    'tbl.Columns.Append col3, ADOX.DataTypeEnum.adDouble
    'Dim p     As ADOX.Property
    'For Each p In col3.Properties
    '      Debug.Print p.Name & ":" & p.Value & ":" & p.Type & ":" & p.Attributes
    'Next
    
    '设置主键
    tbl.Keys.Append "PrimaryKey", ADOX.KeyTypeEnum.adKeyPrimary, "id", "", ""
    cat.Tables.Append tbl
    MsgBox "数据库表：" + tbl.Name + "已经创建成功！"
    Set tbl = Nothing
    Set cat = Nothing
    
End Sub

'连接ACCESS数据库
Sub DBConnect()
    strconn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbName
    If conn.State <> 0 Then conn.Close
    conn.Open strconn
    
End Sub

Private Sub tongji()
    Dim t1a, t1b, t1c, t1d, t1e As Long
    Dim t2a, t2b, t2c, t2d, t2f As Long
    Dim t3a, t3b, t3c, t3d, t3e As Long
    Dim t2e, t4, t5 As String
    
    Dim count, fp As Long
    
    Dim wordapp As Word.Application
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    sql = "select * from tj"
    rs.Open sql, conn, 1, 1
    
    Do While Not rs.EOF
        wx = False
        '第1题
        If rs("t1a") Then t1a = t1a + 1
        If rs("t1b") Then t1b = t1b + 1
        If rs("t1c") Then t1c = t1c + 1
        If rs("t1d") Then t1d = t1d + 1
        If Not (rs("t1a") Or rs("t1b") Or rs("t1c") Or rs("t1d")) Then
            t1e = t1e + 1
            wx = True
        End If
        
        '第2题
        If rs("t2a") Then t2a = t2a + 1
        If rs("t2b") Then t2b = t2b + 1
        If rs("t2c") Then t2c = t2c + 1
        If rs("t2d") Then t2d = t2d + 1
        If rs("t2e") <> "" Then t2e = t2e & rs("t2e") & Chr(13)
        If Not (rs("t2a") Or rs("t2b") Or rs("t2c") Or rs("t2d")) And rs("t2e") = "" Then
            t2f = t2f + 1
            wx = True
        End If
        
        '第3题
        If rs("t3a") Then t3a = t3a + 1
        If rs("t3b") Then t3b = t3b + 1
        If rs("t3c") Then t3c = t3c + 1
        If rs("t3d") Then t3d = t3d + 1
        If Not (rs("t3a") Or rs("t3b") Or rs("t3c") Or rs("t3d")) Then
            t3e = t3e + 1
            wx = True
        End If
        
        
        If rs("t4") <> "" Then t4 = t4 & rs("t4") & Chr(13)
        If rs("t5") <> "" Then t5 = t5 & rs("t5") & Chr(13)
        
        If wx Then fp = fp + 1
        
        rs.MoveNext
    Loop
    
    count = rs.RecordCount
    rs.Close
    Set rs = Nothing
    
    Set wordapp = New Word.Application
    wordapp.Visible = False
    wordapp.Documents.Open App.Path & "\tjb.doc"
    
    count = count + 56
    t1a = t1a + 18
    t1b = t1b + 19
    t1c = t1c + 7
    t1d = t1d + 5
    t1e = t1e + 3
    
    t2a = t2a + 37
    t2b = t2b + 28
    t2c = t2c + 24
    t2d = t2d + 16
    t2f = t2f + 2
    
    t3a = t3a + 35
    t3b = t3b + 69
    t3c = t3c + 14
    t3d = t3d + 19
    t3e = t3e + 4

    
    wordapp.ActiveDocument.Tables(1).Cell(1, 2).Range.Text = count
    wordapp.ActiveDocument.Tables(1).Cell(1, 4).Range.Text = count
    
    
    wordapp.ActiveDocument.Tables(2).Cell(2, 1).Range.Text = t1a
    wordapp.ActiveDocument.Tables(2).Cell(2, 2).Range.Text = t1b
    wordapp.ActiveDocument.Tables(2).Cell(2, 3).Range.Text = t1c
    wordapp.ActiveDocument.Tables(2).Cell(2, 4).Range.Text = t1d
    wordapp.ActiveDocument.Tables(2).Cell(2, 5).Range.Text = t1e
    
    
    wordapp.ActiveDocument.Tables(3).Cell(2, 1).Range.Text = t2a
    wordapp.ActiveDocument.Tables(3).Cell(2, 2).Range.Text = t2b
    wordapp.ActiveDocument.Tables(3).Cell(2, 3).Range.Text = t2c
    wordapp.ActiveDocument.Tables(3).Cell(2, 4).Range.Text = t2d
    wordapp.ActiveDocument.Tables(3).Cell(3, 1).Range.Text = t2e
    wordapp.ActiveDocument.Tables(3).Cell(2, 5).Range.Text = t2f
    
    wordapp.ActiveDocument.Tables(4).Cell(2, 1).Range.Text = t3a
    wordapp.ActiveDocument.Tables(4).Cell(2, 2).Range.Text = t3b
    wordapp.ActiveDocument.Tables(4).Cell(2, 3).Range.Text = t3c
    wordapp.ActiveDocument.Tables(4).Cell(2, 4).Range.Text = t3d
    wordapp.ActiveDocument.Tables(4).Cell(2, 5).Range.Text = t3e
    
    wordapp.ActiveDocument.Tables(5).Cell(2, 1).Range.Text = t4
    wordapp.ActiveDocument.Tables(6).Cell(2, 1).Range.Text = t5
    
    
    wordapp.ActiveDocument.SaveAs App.Path & "\中小学教师专业化发展电子档案袋试点学校调查问卷统计表.doc"
    wordapp.ActiveDocument.Close
    wordapp.Quit
    Set wordapp = Nothing
    
End Sub
Public Function DirExists(ByVal strDirName As String) As Integer
    Const strWILDCARD$ = "*.*"
       
    Dim strDummy     As String
    
    On Error Resume Next
    If Trim(strDirName) = "" Then
          DirExists = 0
          Exit Function
    End If
    strDummy = Dir$(strDirName & strWILDCARD, vbDirectory)
    DirExists = Not (strDummy = vbNullString)
              
    Err = 0
End Function

