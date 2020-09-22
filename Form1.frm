VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7410
   LinkTopic       =   "Form1"
   ScaleHeight     =   6195
   ScaleWidth      =   7410
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   615
      Left            =   2160
      TabIndex        =   2
      Top             =   4680
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   3600
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   5760
      Width           =   7095
   End
   Begin VB.Menu mnubookmark 
      Caption         =   "Bookmark"
      Begin VB.Menu mnufavs 
         Caption         =   "-"
         Index           =   0
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const CSIDL_FAVORITES = 6
Private Declare Sub DoOrganizeFavDlg Lib "shdocvw.dll" (ByVal hwnd As Long, ByVal path As String)
Private Declare Function SHGetSpecialFolderPath Lib "shell32.dll" Alias "SHGetSpecialFolderPathA" _
    (ByVal hwnd As Long, ByVal sPath As String, _
    ByVal Folder As Long, ByVal Create As Long) As Long
Private Type tFaves
  Title As String
  URL  As String
End Type
Dim sEntry As String
Dim ncnt1 As Integer
Private Faves() As tFaves
Dim intConst            '*' Integer for loop construction
Dim intCounter          '*' Integer for item counter

Sub ReadFavorites(sEntry As String)
  On Error Resume Next
     
  Static nIndent As Integer  ' used for formatting output
  Static ncnt1  As Integer
  Dim sDir As String
  Dim aSubs() As String
  Dim SubCount As Integer
  Dim X As Integer
  Dim cp
  Dim sURL As String
  Dim sFile As String
  Dim sTemplate As String
  Dim Search, Where   ' Declare variables.
ncnt1 = 1
  ' increment indentation count
  nIndent = nIndent + 1
     
  ReDim aSubs(0) As String
     
  sDir = sEntry
  If Right$(sEntry, 1) <> "\" Then
    sEntry = sEntry & "\"
  End If
     
  sTemplate = "*.*"
     
  ChDir sEntry
     

  ' Get Directories
  sEntry = Dir$(sDir & "\*.*", vbDirectory)
  If Err Then Exit Sub
  Do While Len(sEntry)
    If (GetAttr(sDir & "\" & sEntry) And vbDirectory) <> 0 Then
      SubCount = SubCount + 1
      ReDim Preserve aSubs(SubCount) As String
      aSubs(SubCount) = sDir & "\" & sEntry
    End If
    sEntry = Dir$
  Loop
     
  ' Output Directory
  'Debug.Print Space((nIndent - 1) * 3) & "->" & StripPath(sDir)
  'Text1.SelText = Space((nIndent - 1) * 3) & "->" & StripPath(sDir)
  ' Get Files
  sEntry = Dir$(sDir & "\*.*", vbDirectory)
  Do While Len(sEntry)
    sFile = sDir & "\" & sEntry
    If (GetAttr(sFile) And vbDirectory) = 0 Then
      Open sFile For Input Access Read As 1
      sURL = Mid(Input(FileLen(sFile), 1), 21)
      sURL = Left(sURL, Len(sURL) - 2)
      If InStr(sURL, Chr$(13)) > 0 Then
        sURL = Left$(sURL, InStr(sURL, Chr$(13)) - 1)
      End If
      Close #1
      ' Output file/URL
      If Trim$(sURL) <> "" Then
        ReDim Preserve Faves(0 To ncnt1)
        Faves(ncnt1).Title = Left(sEntry, Len(sEntry) - 4)
        Faves(ncnt1).URL = Mid$(sURL, 5)
 'intCounter = 1
 Load mnufavs(ncnt1)
  mnufavs(ncnt1).Caption = Faves(ncnt1).Title
        ncnt1 = ncnt1 + 1
        'intCounter = intCounter + 1
      End If
    End If
    sEntry = Dir$
  Loop
     
  ' Recurse through sub-directories
  For X = 1 To SubCount
    If Right(aSubs(X), 1) <> "." Then
      ReadFavorites aSubs(X)
    End If
  Next
     
  ' decrement indentation count
  nIndent = nIndent - 1
 
End Sub
Function StripPath(sPath As String) As String

         Dim n As Integer
         Dim nPos As Integer
            
            
         n = 1
         While n > 0
            nPos = n
            n = InStr(n + 1, sPath, "\")
         Wend
            
         If nPos > 0 Then
            StripPath = Mid(sPath, nPos + 1)
         Else
            StripPath = sPath
         End If
            

      End Function

Private Sub Command1_Click()

ReadFavorites (sEntry)

End Sub

Private Sub Command2_Click()
'On Error Resume Next

Dim mCount As Integer
Do While mnufavs.Count > 1

mCount = mnufavs.Count
'MsgBox mCount & "   " & " / " & mnufavs.Count
Unload mnufavs(mCount - 1)
Loop
End Sub

Private Sub Form_Load()
     Dim path As String
  
  

path = String$(260, 0)
    SHGetSpecialFolderPath Me.hwnd, path, CSIDL_FAVORITES, 1
 Text1 = path
sEntry = Text1
End Sub

Private Sub mnufavs_Click(Index As Integer)

On Error Resume Next
'MsgBox mnufavs(Index).Caption


MsgBox mnufavs(Index).Index & "   /   " & Faves(Index).URL
End Sub
