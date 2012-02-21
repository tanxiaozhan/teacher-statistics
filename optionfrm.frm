VERSION 5.00
Begin VB.Form optionfrm 
   Caption         =   "选项设置"
   ClientHeight    =   2190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6195
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2190
   ScaleWidth      =   6195
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton Command3 
      Caption         =   "确    定"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3345
      TabIndex        =   4
      Top             =   1380
      Width           =   1440
   End
   Begin VB.CommandButton Command2 
      Caption         =   "恢复缺省值"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   1200
      TabIndex        =   3
      Top             =   1380
      Width           =   1440
   End
   Begin VB.CommandButton Command1 
      Caption         =   "浏览"
      Height          =   375
      Left            =   5430
      TabIndex        =   2
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   720
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   615
      Width           =   4710
   End
   Begin VB.Label Label1 
      Caption         =   "路径"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   195
      TabIndex        =   1
      Top             =   675
      Width           =   480
   End
End
Attribute VB_Name = "optionfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim BI     As BROWSEINFO
  Dim NFolder     As Long
  Dim IDL     As ITEMIDLIST
  Dim PIdl     As Long
  Dim SPath     As String
  Dim SHFI     As SHFILEINFO
  Dim M_wCurOptIdx     As Integer
  Dim TxtPath     As String
  Dim TxtDisplayName     As String
  Dim Noerror     As Boolean
  Dim SHGFI_PIDL     As Long
  Dim Shgfi_Icon     As Long
  Dim Shgfi_Smallicon     As Long
    
  With BI
          .hOwner = Me.hwnd
          NFolder = GetFolderValue(M_wCurOptIdx)
            
          If SHGetSpecialFolderLocation(ByVal Me.hwnd, ByVal NFolder, IDL) = Noerror Then
              .pidlRoot = IDL.mkid.cb
          End If
            
          .pszDisplayName = String$(MAX_PATH, 0)
          .lpszTitle = "请选择调查问卷所在的文件夹。"
          .ulFlags = 0
  End With
        
  TxtPath = ""
  TxtDisplayName = ""
    
  PIdl = SHBrowseForFolder(BI)
        
  If PIdl = 0 Then Exit Sub
  SPath = String$(MAX_PATH, 0)
  SHGetPathFromIDList ByVal PIdl, ByVal SPath
    
  TxtPath = Left(SPath, InStr(SPath, vbNullChar) - 1)
  TxtDisplayName = Left$(BI.pszDisplayName, InStr(BI.pszDisplayName, vbNullChar) - 1)
  SHGetFileInfo ByVal PIdl, 0&, SHFI, Len(SHFI), SHGFI_PIDL Or Shgfi_Icon Or Shgfi_Smallicon
  SHGetFileInfo ByVal PIdl, 0&, SHFI, Len(SHFI), SHGFI_PIDL Or Shgfi_Icon
  CoTaskMemFree PIdl
  Text1.Text = TxtPath
  'txtpath就是目录所在的路径   End Sub
  
End Sub
Private Function GetFolderValue(wIdx As Integer) As Long
    If wIdx < 2 Then
        GetFolderValue = 0
    ElseIf wIdx < 12 Then
        GetFolderValue = wIdx
    Else
        GetFolderValue = wIdx + 4
    End If
End Function
Private Sub Command2_Click()
    Text1.Text = App.Path & "\doc"
End Sub

Private Sub Command3_Click()
    docPath = Text1.Text
    Unload Me
End Sub

Private Sub Form_Load()
    Text1.Text = docPath
End Sub
