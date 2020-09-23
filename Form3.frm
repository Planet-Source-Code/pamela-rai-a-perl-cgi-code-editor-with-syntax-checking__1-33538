VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Open file for..."
   ClientHeight    =   1980
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2850
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   2850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   1560
      Width           =   1095
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Append"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Writing"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Reading"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Value           =   -1  'True
      Width           =   975
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim rwa As String
If Option1.Value = True Then
rwa = "<"
End If
If Option2.Value = True Then
rwa = ">"
End If
If Option3.Value = True Then
rwa = ">>"
End If

xxx = "open (INFIL, """ + rwa + "/htdocs/"");" + _
vbCrLf + "@data = <INFIL>;" + _
vbCrLf + "close INFIL;" + vbCrLf
Form1.R1.SelText = xxx

Form3.Hide
'Form1.Update_Click
End Sub

Private Sub Form_Load()
SetWindowPos Form3.hWnd, -1, 0, 0, 0, 0, 1 Or 2
End Sub
