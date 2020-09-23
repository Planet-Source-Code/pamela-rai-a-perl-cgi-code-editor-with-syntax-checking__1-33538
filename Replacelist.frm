VERSION 5.00
Begin VB.Form Replacelist 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Replacelist"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10695
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   10695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox List3 
      Height          =   1425
      Left            =   9720
      TabIndex        =   7
      Top             =   1560
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Remove 
      Caption         =   "Remove >>"
      Height          =   255
      Left            =   9600
      TabIndex        =   6
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   0
      TabIndex        =   5
      Top             =   360
      Width           =   9495
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   480
      Top             =   3360
   End
   Begin VB.ListBox List2 
      Height          =   2595
      ItemData        =   "Replacelist.frx":0000
      Left            =   0
      List            =   "Replacelist.frx":0002
      MultiSelect     =   2  'Extended
      TabIndex        =   4
      Top             =   720
      Width           =   9495
   End
   Begin VB.CommandButton Lukk 
      Caption         =   "Lukk"
      Height          =   315
      Left            =   9600
      TabIndex        =   3
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton Add 
      Caption         =   "<< Add"
      Height          =   255
      Left            =   9600
      TabIndex        =   2
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10695
   End
   Begin VB.ListBox List1 
      Height          =   2595
      ItemData        =   "Replacelist.frx":0004
      Left            =   0
      List            =   "Replacelist.frx":0006
      TabIndex        =   0
      Top             =   3360
      Width           =   9495
   End
End
Attribute VB_Name = "Replacelist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Add_Click()
List1.AddItem Text2.Text, 0
List2.AddItem Text2.Text, 0
End Sub

Private Sub Form_Load()
SetWindowPos Replacelist.hwnd, -1, 0, 0, 0, 0, 1 Or 2


End Sub

Private Sub List1_Click()
If active = False Then
List2.ListIndex = List1.ListIndex
Text1.Text = List2.List(List2.ListIndex)
End If
End Sub

Private Sub List1_DblClick()
Text1.Text = List1.List(List1.ListIndex)
End Sub

Private Sub List2_Click()
If active = False Then
List1.ListIndex = List2.ListIndex
Text1.Text = List2.List(List2.ListIndex)
End If
End Sub

Private Sub List2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
Remove_Click
End If
End Sub

Private Sub Lukk_Click()
Replacelist.Hide
Batchfiles.Show
End Sub

Private Sub Remove_Click()
List3.Clear
For i = 0 To List2.ListCount - 1

If List2.Selected(i) = True Then
List3.AddItem i, 0
'List1.RemoveItem List1.List(i)
'List2.RemoveItem List2.List(i)
End If
Next
For i = 0 To List3.ListCount - 1
List2.RemoveItem List3.List(i)
List1.RemoveItem List3.List(i)
Next
End Sub

Private Sub Text1_Change()
If Not List2.ListIndex = -1 Then
List2.List(List2.ListIndex) = Text1.Text
End If
End Sub
