VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Find"
   ClientHeight    =   1485
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5820
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   5820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command4 
      Caption         =   "Find Next"
      Height          =   255
      Left            =   3120
      TabIndex        =   5
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Replace All"
      Height          =   255
      Left            =   1560
      TabIndex        =   4
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Replace"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   5535
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Find"
      Height          =   255
      Left            =   4560
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
'MsgBox Form1.R1.SelText
returnval = Form1.R1.Find(Text1.Text, Form1.R1.SelStart)
Form1.R1.SetFocus
If returnval = -1 Then
Form4.Hide
commandb = MsgBox("End of document! Start over?", vbYesNo)
End If
If commandb = 6 Then
Form1.R1.SelStart = 1
Form1.R1.SetFocus
Form4.Show
End If
If commandb = 7 Then
Form4.Hide
End If
End Sub

Private Sub Command2_Click()
If Len(Form1.R1.SelText) > 0 Then
Form1.R1.SelText = Text2.Text
End If
End Sub

Private Sub Command3_Click()
returnval = Form1.R1.Find(Form4.Text1.Text)
Form1.R1.SetFocus
Form1.R1.SelText = Text2.Text
Do While Not returnval = -1
returnval = Form1.R1.Find(Form4.Text1.Text, returnval + Len(Text1.Text))
Form1.R1.SelText = Text2.Text
Loop
End Sub

Private Sub Command4_Click()
returnval = Form1.R1.Find(Form4.Text1.Text, Form1.R1.SelStart + 1)
Form1.R1.SetFocus
If returnval = -1 Then
Form4.Hide
commandb = MsgBox("End of document! Start over?", vbYesNo)
End If
If commandb = 6 Then
Form1.R1.SelStart = 1
Form1.R1.SetFocus
Form4.Show
End If
If commandb = 7 Then
Form4.Hide
End If
End Sub

Private Sub Form_Load()
SetWindowPos Form4.hwnd, -1, 0, 0, 0, 0, 1 Or 2
Text1.Text = Form1.R1.SelText
End Sub
