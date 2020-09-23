VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   0  'None
   Caption         =   "Select paste text"
   ClientHeight    =   2175
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7110
   LinkTopic       =   "Form5"
   ScaleHeight     =   2175
   ScaleWidth      =   7110
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      Height          =   255
      Left            =   5160
      TabIndex        =   2
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   255
      Left            =   6000
      TabIndex        =   1
      Top             =   1800
      Width           =   735
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   2175
      ItemData        =   "Form5.frx":0000
      Left            =   0
      List            =   "Form5.frx":0002
      TabIndex        =   0
      Top             =   0
      Width           =   7095
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form5.Hide
End Sub

Private Sub Command2_Click()
List1.Clear
End Sub

Private Sub Form_Load()
SetWindowPos Form5.hwnd, -1, 0, 0, 0, 0, 1 Or 2


End Sub

Private Sub Form_Resize()
List1.Width = Form5.Width
List1.Height = Form5.Height
End Sub

Private Sub List1_Click()
Form1.R1.SelText = List1.List(List1.ListIndex)
Form5.Hide
End Sub
