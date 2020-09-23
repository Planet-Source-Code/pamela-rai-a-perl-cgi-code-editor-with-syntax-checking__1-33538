VERSION 5.00
Begin VB.Form Vars 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Vars"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   1830
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   1830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox List2 
      Height          =   2985
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1815
   End
   Begin VB.ListBox List1 
      Height          =   2985
      ItemData        =   "Vars.frx":0000
      Left            =   3600
      List            =   "Vars.frx":0002
      MultiSelect     =   1  'Simple
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   -120
      Width           =   1815
   End
End
Attribute VB_Name = "Vars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
SetWindowPos Vars.hWnd, -1, 0, 0, 0, 0, 1 Or 2

End Sub


Private Sub List2_Click()
Form1.R1.SelText = List2.List(List2.ListIndex)
End Sub
