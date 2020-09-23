VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Batchfiles 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Batch files"
   ClientHeight    =   3795
   ClientLeft      =   3750
   ClientTop       =   105
   ClientWidth     =   10815
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   3795
   ScaleWidth      =   10815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.FileListBox File1 
      Height          =   1260
      Left            =   7920
      MultiSelect     =   2  'Extended
      TabIndex        =   10
      Top             =   2160
      Width           =   2895
   End
   Begin VB.DirListBox Dir1 
      Height          =   1665
      Left            =   7920
      TabIndex        =   9
      Top             =   360
      Width           =   2895
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   7920
      TabIndex        =   8
      Top             =   0
      Width           =   2895
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1200
      Top             =   1080
   End
   Begin MSComctlLib.ProgressBar P1 
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   3600
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Min             =   1e-4
   End
   Begin VB.CommandButton Extractlinks 
      Caption         =   "Extract links"
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Top             =   3120
      Width           =   1695
   End
   Begin VB.CommandButton Lukk 
      Caption         =   "Lukk"
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton Editreplacelist 
      Caption         =   "Edit Replace List"
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   3120
      Width           =   1695
   End
   Begin VB.CommandButton Remove 
      Caption         =   "Remove"
      Height          =   375
      Left            =   4680
      TabIndex        =   3
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton Addfile 
      Caption         =   "Add File(s).."
      Height          =   375
      Left            =   6840
      TabIndex        =   2
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton runbatch 
      Caption         =   "Run Batch"
      Height          =   375
      Left            =   5760
      TabIndex        =   1
      Top             =   3120
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   2985
      ItemData        =   "Batchfiles.frx":0000
      Left            =   0
      List            =   "Batchfiles.frx":0002
      OLEDropMode     =   1  'Manual
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   7815
   End
End
Attribute VB_Name = "Batchfiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'
' Desktop—virtual folder
Const CSIDL_DESKTOP = &H0
'
' User's program groups
Const CSIDL_PROGRAMS = &H2
'
' Control Panel.
Const CSIDL_CONTROLS = &H3
'
' Folder containing installed printers.
Const CSIDL_PRINTERS = &H4
'
' Folder that serves as a common repository for documents.
Const CSIDL_PERSONAL = &H5
'
' Folder that serves as a common repository for the user's favorite items.
Const CSIDL_FAVORITES = &H6
'
' Folder that corresponds to the user's Startup program group.
Const CSIDL_STARTUP = &H7
'
' User's most recently used documents.
Const CSIDL_RECENT = &H8
'
' Folder that contains Send To menu items.
Const CSIDL_SENDTO = &H9
'
' Recycle Bin.
Const CSIDL_BITBUCKET = &HA
'
' Start menu items.
Const CSIDL_STARTMENU = &HB
'
' Folder used to physically store file objects on the desktop.
Const CSIDL_DESKTOPDIRECTORY = &H10
'
' My Computer—virtual folder
Const CSIDL_DRIVES = &H11
'

Const FO_MOVE = &H1
Const FO_RENAME = &H4
Const FOF_SILENT = &H4
Const FOF_NOCONFIRMATION = &H10
Const FOF_RENAMEONCOLLISION = &H8
Const MAX_PATH As Integer = 260
Const SHARD_PATH = &H2&
Const SHCNF_IDLIST = &H0
Const SHCNE_ALLEVENTS = &H7FFFFFFF

Private Type SHFILEOPSTRUCT
    hwnd      As Long
    wFunc     As Long
    pFrom     As String
    pTo       As String
    fFlags    As Integer
    fAborted  As Boolean
    hNameMaps As Long
    sProgress As String
End Type

Private Type SHITEMID
    cb   As Long
    abID As Byte
End Type

Private Type ITEMIDLIST
    mkid As SHITEMID
End Type




Private Declare Function SHGetSpecialFolderLocation Lib "Shell32.dll" _
    (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function SHGetSpecialFolderLocationD Lib "Shell32.dll" _
    Alias "SHGetSpecialFolderLocation" (ByVal hwndOwner As Long, _
    ByVal nFolder As Long, ByRef ppidl As Long) As Long
Private Declare Function SHGetPathFromIDList Lib "Shell32.dll" _
    Alias "SHGetPathFromIDListA" (ByVal pidl As Long, _
    ByVal pszPath As String) As Long


Private Sub Addfile_Click()
Dim xret
For i = 0 To File1.ListCount - 1
If File1.Selected(i) = True Then
Batchfiles.List1.AddItem Dir1.Path + "\" + File1.List(i)
End If
Next






End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Dir1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
Dir1.Path = fGetSpecialFolder(CSIDL_DESKTOPDIRECTORY)
End If
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub Editreplacelist_Click()
Batchfiles.Hide
Replacelist.Show
End Sub

Private Sub Extractlinks_Click()
active = True

P1.Min = 0
P1.Max = (List1.ListCount) * 6
Dim thelink
Dim findpos1
Dim findpos2
Dim startx
Dim i
'Replacelist.Show
For i = 0 To List1.ListCount - 1
Form1.R2.LoadFile List1.List(i), rtfText
P1.Value = P1.Value + 1
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
startx = 1
findpos1 = InStr(startx, Form1.R2.Text, """>/", vbTextCompare)
Do While findpos1 > 0
findpos1 = InStr(startx, Form1.R2.Text, """>/", vbTextCompare)
findpos2 = InStr(findpos1 + 1, Form1.R2.Text, """", vbTextCompare)
startx = findpos2
If findpos1 > 0 And findpos2 > 0 Then
thelink = Mid(Form1.R2.Text, findpos1 + 2, findpos2 - (findpos1 + 2))
Replacelist.List1.AddItem thelink
End If
Loop
P1.Value = P1.Value + 1
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
startx = 1
findpos1 = InStr(startx, Form1.R2.Text, """/", vbTextCompare)
Do While findpos1 > 0
findpos1 = InStr(startx, Form1.R2.Text, """/", vbTextCompare)
findpos2 = InStr(findpos1 + 1, Form1.R2.Text, """", vbTextCompare)
startx = findpos2
If findpos1 > 0 And findpos2 > 0 Then
thelink = Mid(Form1.R2.Text, findpos1 + 1, findpos2 - (findpos1 + 1))
Replacelist.List1.AddItem thelink
End If
Loop
P1.Value = P1.Value + 1
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
startx = 1
findpos1 = InStr(startx, Form1.R2.Text, """</", vbTextCompare)
Do While findpos1 > 0
findpos1 = InStr(startx, Form1.R2.Text, """</", vbTextCompare)
findpos2 = InStr(findpos1 + 1, Form1.R2.Text, """", vbTextCompare)
startx = findpos2
If findpos1 > 0 And findpos2 > 0 Then
thelink = Mid(Form1.R2.Text, findpos1 + 2, findpos2 - (findpos1 + 2))
Replacelist.List1.AddItem thelink
End If
Loop
P1.Value = P1.Value + 1
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
startx = 1
findpos1 = InStr(startx, Form1.R2.Text, """http", vbTextCompare)
Do While findpos1 > 0
findpos1 = InStr(startx, Form1.R2.Text, """http", vbTextCompare)
findpos2 = InStr(findpos1 + 1, Form1.R2.Text, """", vbTextCompare)
startx = findpos2
If findpos1 > 0 And findpos2 > 0 Then
thelink = Mid(Form1.R2.Text, findpos1 + 1, findpos2 - (findpos1 + 1))
Replacelist.List1.AddItem thelink
End If
Loop
P1.Value = P1.Value + 1
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
startx = 1
findpos1 = InStr(startx, Form1.R2.Text, "'/", vbTextCompare)
Do While findpos1 > 0
findpos1 = InStr(startx, Form1.R2.Text, "'/", vbTextCompare)
findpos2 = InStr(findpos1 + 1, Form1.R2.Text, ";", vbTextCompare)
startx = findpos2
If findpos1 > 0 And findpos2 > 0 Then
thelink = Mid(Form1.R2.Text, findpos1 + 1, findpos2 - (findpos1 + 1))
Replacelist.List1.AddItem thelink
End If
Loop
'P1.Value = P1.Value + 1

Next
P1.Value = 0

RemoveDupes Replacelist.List1
For i = 0 To Replacelist.List1.ListCount - 1
Replacelist.List2.List(i) = Replacelist.List1.List(i)
Next

'RemoveDupes Replacelist.List2
Replacelist.Show
active = False
End Sub

Private Sub RemoveDupes(lst As ListBox)
    Dim iPos As Integer
    iPos = 0
    '-- if listbox empty then exit..
    If lst.ListCount < 1 Then Exit Sub


    Do While iPos < lst.ListCount
        lst.Text = lst.List(iPos)
        '-- check if text already exists..


        If lst.ListIndex <> iPos Then
            '-- if so, remove it and keep iPos..
            lst.RemoveItem iPos
        Else
            '-- if not, increase iPos..
            iPos = iPos + 1
        End If
    Loop
    '-- used to unselect the last selected l
    '     ine..
    lst.Text = "~~~^^~~~"
End Sub

Private Sub File1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
Addfile_Click
End If

End Sub

Private Sub Form_Load()
'SetWindowPos Batchfiles.hWnd, -1, 0, 0, 0, 0, 1 Or 2
Dir1.Path = fGetSpecialFolder(CSIDL_DESKTOPDIRECTORY)
File1.Pattern = "*.pl;*.cgi"
End Sub

Private Sub Image1_Click()
Dir1.Path = sGetDesktop
MsgBox sGetDesktop
End Sub

Private Sub List1_DragDrop(Source As Control, X As Single, Y As Single)
Batchfiles.List1.AddItem strOpen
End Sub

Private Sub List1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim xret
Dim i
For i = 1 To Data.Files.Count
xret = InStr(1, Data.Files(i), ".pl", vbTextCompare)
If Not xret = 0 Then
Batchfiles.List1.AddItem Data.Files(i)
Else
MsgBox "Not a valid Perl file (*.pl) !!!" + vbCrLf + Data.Files(i)
End If
Next
End Sub

Private Sub Lukk_Click()
Batchfiles.Hide
End Sub
Private Function fGetSpecialFolder(CSIDL As Long) As String
Dim sPath As String
Dim IDL   As ITEMIDLIST
'
' Retrieve info about system folders such as the
' "Recent Documents" folder.  Info is stored in
' the IDL structure.
'
fGetSpecialFolder = ""
If SHGetSpecialFolderLocation(Me.hwnd, CSIDL, IDL) = 0 Then
    '
    ' Get the path from the ID list, and return the folder.
    '
    sPath = Space$(MAX_PATH)
    If SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal sPath) Then
        fGetSpecialFolder = Left$(sPath, InStr(sPath, vbNullChar) - 1) & "\"
    End If
End If
End Function
Private Sub Remove_Click()
If Not List1.ListIndex = -1 Then
List1.RemoveItem (List1.ListIndex)
End If
End Sub

Private Sub runbatch_Click()
For i = 0 To List1.ListCount - 1
Form1.R2.LoadFile List1.List(i), rtfText
For X = 0 To Replacelist.List1.ListCount - 1
Form1.R2.Text = Replace(Form1.R2.Text, Replacelist.List1.List(X), Replacelist.List2.List(X), 1, -1, vbTextCompare)
Next
Form1.R2.SaveFile List1.List(i), rtfText
Next
MsgBox "Batch ok"
End Sub

Private Sub Timer1_Timer()
If List1.ListCount > 0 Then
Extractlinks.Enabled = True
Remove.Enabled = True
Else
Extractlinks.Enabled = False
Remove.Enabled = False
End If

If Replacelist.List1.ListCount > 0 And List1.ListCount > 0 Then
runbatch.Enabled = True
Else
runbatch.Enabled = False
End If
End Sub
