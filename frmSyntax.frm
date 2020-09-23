VERSION 5.00
Begin VB.Form frmSyntax 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Syntax"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10665
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   10665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   4335
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "frmSyntax.frx":0000
      Top             =   120
      Width           =   10455
   End
End
Attribute VB_Name = "frmSyntax"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Path to download activeperl for windows : http://aspn.activestate.com/ASPN/Downloads/ActivePerl/
'This will report errors in you perlcode by klicking syntax
perlpath = "F:\Perl\bin\perl.exe" 'link to activeperl



Private Sub Form_Load()
On Error GoTo errhandle
Dim texten As String
'MsgBox Form1.dlgOpenFile.FileName
texten = Form1.dlgOpenFile.FileName

texten = GetShortFileName(texten)


'MsgBox texten
Text1.Text = GetCommandOutput(perlpath & " " & texten, True, True, True)
Text1.Text = Replace(Text1.Text, vbLf, vbCrLf)
Exit Sub
errhandle:

End Sub

