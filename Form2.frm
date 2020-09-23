VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Split data"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   1560
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   1560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   2040
      Width           =   855
   End
   Begin VB.TextBox varx 
      Height          =   285
      Left            =   240
      TabIndex        =   5
      Text            =   "$i"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox charx 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Text            =   ";"
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox numberx 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "2"
      Top             =   360
      Width           =   615
   End
   Begin VB.Label label3 
      Caption         =   "Var to split:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Split char:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Number of values:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim aaa As String
Dim xxx As String
Private Sub Command1_Click()
xxx = "("
For i = 1 To Form2.numberx.Text
aaa = i
xxx = xxx + "$var" + aaa + ", "
Next
xxx = xxx + ") = split(/" + Form2.charx.Text + _
"/, " + Form2.varx.Text + ");" + vbCrLf
xxx = Replace(xxx, ", )", ")")
Clipboard.SetText xxx
Form1.R1.SelText = Clipboard.GetText
Form2.Hide

'Form1.Update_Click
End Sub

Private Sub Form_Load()
SetWindowPos Form2.hWnd, -1, 0, 0, 0, 0, 1 Or 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
Close
End Sub

