Attribute VB_Name = "Module1"
'*******************************浏览文件夹*****************************************************
  Public Declare Function SHBrowseForFolder Lib "shell32.dll" Alias _
                  "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
  
  Public Declare Function SHGetSpecialFolderLocation Lib _
                  "shell32.dll" (ByVal hwndOwner As Long, ByVal NFolder _
                  As Long, PIdl As ITEMIDLIST) As Long
    
  Public Declare Function SHGetFileInfo Lib "Shell32" Alias _
                  "SHGetFileInfoA" (ByVal pszPath As Any, ByVal _
                  dwFileAttributes As Long, psfi As SHFILEINFO, ByVal _
                  cbFileInfo As Long, ByVal uFlags As Long) As Long
    
  Public Declare Function ShellAbout Lib "shell32.dll" Alias _
                  "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As _
                  String, ByVal szOtherStuff As String, ByVal hIcon As Long) _
                  As Long
  Public Declare Function SHGetPathFromIDList Lib "shell32.dll" _
                  Alias "SHGetPathFromIDListA" (ByVal PIdl As Long, ByVal _
                  pszPath As String) As Long
  Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
    
  Public Const MAX_PATH = 260
    
  Public Type SHITEMID
          cb   As Long
          abID()   As Byte
  End Type
    
  Public Type ITEMIDLIST
          mkid   As SHITEMID
  End Type
    
  Public Type BROWSEINFO
          hOwner   As Long
          pidlRoot   As Long
          pszDisplayName   As String
          lpszTitle   As String
          ulFlags   As Long
          lpfn   As Long
          lParam   As Long
          iImage   As Long
  End Type
    
  Public Type SHFILEINFO
          hIcon   As Long
          iIcon   As Long
          dwAttributes   As Long
          szDisplayName   As String * MAX_PATH
          szTypeName   As String * 80
  End Type
  '*******************************浏览文件夹*****************************************************

'************全局变量*****************
Public dbName As String
Public conn As ADODB.Connection
Public docPath As String

