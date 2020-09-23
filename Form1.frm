VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5565
   ClientLeft      =   3765
   ClientTop       =   3975
   ClientWidth     =   7575
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5565
   ScaleWidth      =   7575
   Begin RichTextLib.RichTextBox R2 
      Height          =   1815
      Left            =   4800
      TabIndex        =   5
      Top             =   3360
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   3201
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Form1.frx":0442
   End
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   480
      Top             =   1920
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   6240
      TabIndex        =   4
      Text            =   "0"
      Top             =   1200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   6240
      TabIndex        =   3
      Text            =   "0"
      Top             =   600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   6000
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   480
      Top             =   1440
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   5310
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            AutoSize        =   2
            Enabled         =   0   'False
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgOpenFile 
      Left            =   480
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox R1 
      Height          =   5535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   9763
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   3
      DisableNoScroll =   -1  'True
      RightMargin     =   50000
      TextRTF         =   $"Form1.frx":052E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu New 
         Caption         =   "New"
      End
      Begin VB.Menu Open 
         Caption         =   "Open"
      End
      Begin VB.Menu Save 
         Caption         =   "Save"
      End
      Begin VB.Menu SaveAs 
         Caption         =   "Save As"
      End
   End
   Begin VB.Menu Edit 
      Caption         =   "Edit"
      Begin VB.Menu Copyprivate 
         Caption         =   "Copy private"
      End
      Begin VB.Menu Pasteprivate 
         Caption         =   "Paste private"
      End
      Begin VB.Menu Copy 
         Caption         =   "Copy"
      End
      Begin VB.Menu Paste 
         Caption         =   "Paste"
      End
      Begin VB.Menu Cut 
         Caption         =   "Cut"
      End
      Begin VB.Menu Undo 
         Caption         =   "Undo"
      End
      Begin VB.Menu Selectall 
         Caption         =   "Select All"
      End
      Begin VB.Menu Find 
         Caption         =   "Find"
      End
      Begin VB.Menu Addseltexttobatch 
         Caption         =   "Add seltext to batch replacelist"
      End
      Begin VB.Menu separ1 
         Caption         =   "-"
      End
      Begin VB.Menu Curseltext 
         Caption         =   "---> Current seltext: "
      End
   End
   Begin VB.Menu Insert 
      Caption         =   "Insert"
      Begin VB.Menu printheader 
         Caption         =   "Print Header -- Content ty..."
      End
      Begin VB.Menu Start 
         Caption         =   "Start"
      End
      Begin VB.Menu Readinputlink 
         Caption         =   "Read input link"
      End
      Begin VB.Menu Readinputfromform 
         Caption         =   "Read input from form"
      End
      Begin VB.Menu Foreach 
         Caption         =   "Foreach"
      End
      Begin VB.Menu Splitx 
         Caption         =   "Split"
      End
      Begin VB.Menu Openfile 
         Caption         =   "Open file"
      End
      Begin VB.Menu addarrycell 
         Caption         =   "Add array cell"
      End
      Begin VB.Menu arraysize 
         Caption         =   "Array size"
      End
      Begin VB.Menu if 
         Caption         =   "if"
      End
      Begin VB.Menu ifelse 
         Caption         =   "if else"
      End
      Begin VB.Menu aortarray 
         Caption         =   "Sort array"
      End
   End
   Begin VB.Menu Convert 
      Caption         =   "Convert"
      Begin VB.Menu ToPerl 
         Caption         =   "To Perl"
         Shortcut        =   ^P
      End
      Begin VB.Menu ToHTML 
         Caption         =   "To HTML"
         Shortcut        =   ^H
      End
   End
   Begin VB.Menu Lock 
      Caption         =   "Lock"
      Begin VB.Menu Setlockatcurrentline 
         Caption         =   "Set lock at current line"
      End
      Begin VB.Menu Gotolockedline 
         Caption         =   "Go to locked line"
      End
   End
   Begin VB.Menu Codecolor 
      Caption         =   "Code color"
      Begin VB.Menu Update 
         Caption         =   "Update"
      End
      Begin VB.Menu Reset 
         Caption         =   "Reset"
      End
      Begin VB.Menu codecolen 
         Caption         =   "Codecolor Enabled"
      End
   End
   Begin VB.Menu View 
      Caption         =   "View"
      Begin VB.Menu Variables 
         Caption         =   "Variables"
      End
   End
   Begin VB.Menu Batch 
      Caption         =   "Batch"
      Begin VB.Menu Replaceinselectedfiles 
         Caption         =   "Replace in selected files"
      End
   End
   Begin VB.Menu Syntax 
      Caption         =   "Syntax"
      Begin VB.Menu Checksyntax 
         Caption         =   "Check Syntax"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim saveit As Boolean
Dim xxx As String
Dim resp As String
Dim gblnIgnoreChange As Boolean
Dim gintIndex As Integer
Dim gstrStack(1000) As String
Dim selholder
Dim selmid
Dim clicker
Dim selstartx
Dim sellenghtx
Dim strOpen As String
Dim returnval As Integer
Dim thechar As String
Dim textholder As String
Dim linelock As String










Public Sub GotoLine(LineNum, Highlight As Boolean)
    On Error GoTo done:
  Dim temp As Integer
  Dim Num As Integer
  Dim Pos  As Integer
  Dim LastPos As Integer
  Dim Cut As Integer

    If LineNum = 0 Then Exit Sub
    Pos = 1
    Num = 1
    temp = 0
    Do
        LastPos = temp
        temp = InStr(Pos, R1.Text, vbCrLf)
        If temp = 0 Then GoTo Redo:
        If temp >= 1 Then
            Num = Num + 1
            Pos = temp + 2
        End If
    
    Loop Until Num >= LineNum

    Cut = 1

Redo:
    If temp = 0 Then
        LastPos = 0
        temp = Len(R1.Text)
        Cut = 0
    End If

    If LineNum = 1 Then
        temp = 0
        LastPos = InStr(1, R1.Text, vbCrLf)
        If LastPos = 0 Then
            LastPos = Len(R1.Text)
        End If

        Cut = 0
    End If

    R1.SelStart = temp
    If Highlight = True Then R1.SelLength = LastPos - Cut
    
   R1.SetFocus
done:

End Sub

Private Sub addarrycell_Click()
R1.SelText = "$arrayname[++$#arrayname] = $var;" & vbCrLf
End Sub

Private Sub Addseltexttobatch_Click()
If Len(selholder) > 0 Then
Replacelist.List1.AddItem selholder, 0
Replacelist.List2.AddItem selholder, 0
End If
End Sub





Private Sub aortarray_Click()
R1.SelText = "@array = sort @array;" & vbCrLf
End Sub

Private Sub arraysize_Click()
R1.SelText = "$var = @array;" & vbCrLf
End Sub

Private Sub Checksyntax_Click()
frmSyntax.Show
End Sub

Private Sub codecolen_Click()



If codecolen.Checked = True Then
codecolen.Checked = False

Reset_Click
Else
codecolen.Checked = True
Update_Click
End If
End Sub



Private Sub Copy_Click()
    Clipboard.Clear
    Clipboard.SetText R1.SelText
    R1.SetFocus

End Sub

Private Sub Copyprivate_Click()
'textholder = R1.SelText
Form5.List1.AddItem R1.SelText, 0
End Sub

Private Sub Cut_Click()
    Clipboard.Clear
    Clipboard.SetText R1.SelText
    R1.SelText = ""
    R1.SetFocus
End Sub

Private Sub Find_Click()
Form4.Show
End Sub

Private Sub Foreach_Click()
resp = InputBox("insert Array to use: e.g @array", _
"Enter Array", "@array")
If resp = "" Then

Else
xxx = "foreach $i (" + resp + ") {" + vbCrLf + _
"chomp($i);" + vbCrLf + vbCrLf + _
"}" + vbCrLf
Clipboard.SetText xxx
R1.SelText = Clipboard.GetText
End If

End Sub

Private Sub Form_Load()
selstartx = 0
sellenghtx = 0
clicker = 1
Form1.Caption = "Untitled1.txt"
saveit = False
Form_Resize
linelock = 1
Dim i
For i = 0 To 49
starttext(i) = "x11111111111111111111111111"
Next
End Sub

Private Sub Form_Resize()
'Vars.Left = Form1.Width - Vars.Width
If Form1.Width > 120 And Form1.Height > 720 Then

R1.Width = Form1.Width - 120
R1.Height = Form1.Height - 975
End If
If Form1.WindowState = 1 Then
Vars.Hide
Else
If Vars.List1.ListCount > 0 Then
Vars.Show
End If
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If saveit = True Then
If strOpen = "" Then
strOpen = "Untitled1.txt"
End If
resp = MsgBox("Save " + strOpen, vbYesNo)
If strOpen = "Untitled1.txt" Then
strOpen = ""
End If
If resp = vbYes Then
If strOpen = "" Then
SaveAs_Click
Else
Save_Click
End If
End If
End If
End
End Sub

Private Sub Gotolockedline_Click()
'Dim i As Long
'R1.SelStart = 0
'R1.SelLength = 0
'R1.SelText = ""
GotoLine linelock + 1, False
'For i = 1 To linelock - 1
'SendKeys "{DOWN}", True
'Next
'SendKeys "{RIGHT}", True
'SendKeys "{LEFT}", True
SendKeys "+{HOME}", True
End Sub



Private Sub if_Click()
R1.SelText = "if (){" & vbCrLf & vbCrLf & "}" & vbCrLf
End Sub

Private Sub ifelse_Click()
R1.SelText = "if (){" & vbCrLf & vbCrLf & "}else{" & vbCrLf & vbCrLf & "}" & vbCrLf
End Sub

Private Sub New_Click()
If saveit = True Then
If strOpen = "" Then
strOpen = "Untitled1.txt"
End If
resp = MsgBox("Save " + strOpen, vbYesNo)
If strOpen = "Untitled1.txt" Then
strOpen = ""
End If
If resp = vbYes Then
If strOpen = "" Then
SaveAs_Click
Else
Save_Click
End If
End If
End If
R1.Text = ""
End Sub

Private Sub Open_Click()
linelock = 1
StatusBar1.Panels(3) = ""
dlgOpenFile.Filter = "*.pl Perl files|*.pl|*.txt Text Files|*.txt|*.* All Files|*.*"
   dlgOpenFile.ShowOpen
   strOpen = dlgOpenFile.FileName
   R1.LoadFile strOpen, rtfText
Form1.Caption = dlgOpenFile.FileTitle + " Path: " + strOpen
saveit = False
Update_Click
End Sub

Private Sub Openfile_Click()

Form3.Show
End Sub

Private Sub Paste_Click()
    'Dim pastepos
    'Dim cliplenght
    'Dim i
    'pastepos = R1.SelStart
    'cliplenght = Len(Clipboard.GetText)
    R1.SelText = Clipboard.GetText
    'R1.SelStart = pastepos
    'Update_Click
'R1.SetFocus
'For i = 1 To cliplenght
'SendKeys "{RIGHT}"
'Next
End Sub



Private Sub Pasteprivate_Click()
'R1.SelText = textholder
Form5.Top = ycoord + Form1.Top + 1000
Form5.Left = xcoord + Form1.Left + 500
Form5.Show
End Sub

Private Sub printheader_Click()
R1.SelText = "print ""Content-Type: text/html; charset=iso-8859-1\n\n"";" & vbCrLf
End Sub

Private Sub R1_Change()
 On Error Resume Next
    saveit = True
    If Not gblnIgnoreChange Then
        gintIndex = gintIndex + 1
       gstrStack(gintIndex) = R1.TextRTF
    End If
'checkcolor




End Sub



Private Sub R1_Click()
'Text4.Text = Len(R1.Text)
Form5.Hide
thechar = R1.SelStart
    Dim currLine As Long

    On Local Error Resume Next
    currLine = SendMessage(R1.hwnd, EM_LINEFROMCHAR, -1&, ByVal 0&) + 1
    'MsgBox Format$(currLine, "##,###")
 StatusBar1.Panels(1).Text = "Line: " + Format$(currLine, "##,###")
Dim currentp
Dim i
Dim foundstr
Dim currchar
currentp = thechar
If R1.SelLength > 0 Then
Exit Sub
End If
R1.SelColor = QBColor(0)
If thechar = 0 Then
Exit Sub
End If

For i = 0 To 25
If thechar - i = 0 Then
Exit For
End If
currchar = Mid(R1.Text, thechar - i, 1)


If currchar = "{" Or _
currchar = " " Or _
currchar = ")" Or _
currchar = "," Or _
currchar = "/" Or _
currchar = "]" Or _
currchar = Chr(34) Or _
currchar = "'" Or _
currchar = "<" Or _
currchar = ";" Or _
currchar = "\" Or _
currchar = "-" Or _
currchar = "." Or _
currchar = ":" Or _
currchar = "(" Or _
currchar = "}" Or _
currchar = "[" Or _
currchar = vbCrLf Or _
currchar = vbLf Or _
currchar = "#" Then

Exit For
End If
Next
Text2.Text = thechar - i

For i = 0 To 25
If thechar - i = 0 Then
Exit For
End If
currchar = Mid(R1.Text, thechar + i, 1)


If currchar = "{" Or _
currchar = " " Or _
currchar = ")" Or _
currchar = "," Or _
currchar = "/" Or _
currchar = "]" Or _
currchar = Chr(34) Or _
currchar = "'" Or _
currchar = "<" Or _
currchar = ";" Or _
currchar = "\" Or _
currchar = "-" Or _
currchar = "." Or _
currchar = ":" Or _
currchar = "(" Or _
currchar = "}" Or _
currchar = "[" Or _
currchar = vbCrLf Or _
currchar = vbLf Or _
currchar = "#" Then

Exit For
End If
Next
Text3.Text = currentp + i
'If Not Text2.Text = Text3.Text Then
R1.SelStart = Text2.Text
R1.SelLength = Text3.Text - Text2.Text
Text1.Text = R1.SelText
'End If
Text1.Text = Replace(Text1.Text, vbCrLf, "", 1, -1, vbTextCompare)
Dim finder
finder = InStr(1, Text1.Text, "$", vbTextCompare)
If finder > 0 Then
R1.SelStart = Text2.Text + (finder - 1)
R1.SelLength = Len(Text1.Text) - finder + clicker
End If

If finder > 0 Then
R1.SelColor = &HC000&

'Unload mnuHello(mnuHello.Count)
Else
R1.SelColor = QBColor(0)
End If
'Text1.Text = Trim(Text1.Text)
If Text1.Text = "open " Then
R1.SelColor = QBColor(9)
End If
If Text1.Text = "open" Then
R1.SelColor = QBColor(9)
End If
If Text1.Text = "print" Then
R1.SelColor = QBColor(9)
End If
If Text1.Text = "print " Then
R1.SelColor = QBColor(9)
End If
If Text1.Text = "close" Then
R1.SelColor = QBColor(9)
End If
If Text1.Text = "close " Then
R1.SelColor = QBColor(9)
End If
If Text1.Text = "if " Then
R1.SelColor = QBColor(9)
End If
If Text1.Text = "if" Then
R1.SelColor = QBColor(9)
End If
If Text1.Text = "else " Then
R1.SelColor = QBColor(9)
End If
If Text1.Text = "else" Then
R1.SelColor = QBColor(9)
End If
If Text1.Text = "elsif " Then
R1.SelColor = QBColor(9)
End If
If Text1.Text = "elsif" Then
R1.SelColor = QBColor(9)
End If

'0 Black
'1 Blue
'2 Green
'3 Cyan
'4 Red
'5 Magenta
'6 Yellow
'7 White
'8 Gray
'9 Light Blue
'10 Light Green
'11 Light Cyan
'12 Light Red
'13 Light Magenta
'14 Light Yellow
'15 Bright White
R1.SelLength = 0
R1.SelStart = thechar
R1.SelColor = QBColor(0)
clicker = 1

End Sub



Private Sub Redo_Click()
    gblnIgnoreChange = True
    gintIndex = gintIndex + 1
    On Error Resume Next
    R1.TextRTF = gstrStack(gintIndex)
    gblnIgnoreChange = False
End Sub

Private Sub R1_KeyPress(KeyAscii As Integer)

'Form1.Caption = KeyAscii
If Not KeyAscii And vbCtrlMask Then
'Update_Click

End If

'colorword "open", 0, &H80FF&, KeyAscii
'colorword "close", 1, &HFF8080, KeyAscii

If KeyAscii = 41 Then
R1.SelColor = &H0&
ElseIf KeyAscii = 44 Then
R1.SelColor = &H0&
'Else
'R1.SelColor = &H0&
End If


End Sub

Private Sub R1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 37 Or KeyCode = 38 Or KeyCode = 39 Or KeyCode = 40 Then
clicker = 0
End If
R1_Click

End Sub

Private Sub R1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 And Shift = 0 Then
Me.PopupMenu Edit
End If
If Button = 2 And Shift = 1 Then
Me.PopupMenu Insert
End If
clicker = 0
End Sub

Private Sub R1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

xcoord = X
ycoord = Y
End Sub

Private Sub Readinputfromform_Click()



xxx = "read(STDIN, $buffer, $ENV{""CONTENT_LENGTH""});" + _
vbCrLf + "@pairs = split(/&/, $buffer);" + _
vbCrLf + "foreach $pair (@pairs) {local($name, $value) = split(/=/, $pair);" + _
vbCrLf + "$name =~ tr/+/ /;" + _
vbCrLf + "$name =~ s/%([a-fA-F0-9][a-fA-F0-9])/pack(""C"", hex($1))/eg;" + _
vbCrLf + "$value =~ tr/+/ /;" + _
vbCrLf + "$value =~ s/%([a-fA-F0-9][a-fA-F0-9])/pack(""C"", hex($1))/eg;" + _
vbCrLf + "$co{$name} = $value;" + vbCrLf
R1.SelText = xxx








End Sub

Private Sub Readinputlink_Click()

xxx = "$querystring = $ENV{""QUERY_STRING""};" + _
vbCrLf + "$queryhold = $ENV{""QUERY_STRING""};" + _
vbCrLf + "$querystring =~ s/%([a-fA-F0-9][a-fA-F0-9])/pack(""C"", hex($1))/eg;" + _
vbCrLf + "$querystring =~ tr/+/ /;" + vbCrLf
R1.SelText = xxx

sellenghtx = Len(xxx)



End Sub

Private Sub Replaceinselectedfiles_Click()
Batchfiles.Show


End Sub

Private Sub Reset_Click()
R1.SelStart = 0
R1.SelLength = Len(R1.Text)

R1.SelColor = &H0&
R1.SelStart = 0
R1.SetFocus
End Sub

Private Sub Save_Click()

R1.SaveFile strOpen, rtfText
saveit = False
End Sub

Private Sub SaveAs_Click()
   Dim strNewFile As String
   dlgOpenFile.Filter = "*.pl Perl files|*.pl|*.txt Text Files|*.txt|*.* All Files|*.*"
   dlgOpenFile.ShowSave
   strNewFile = dlgOpenFile.FileName
   R1.SaveFile strNewFile, rtfText
Form1.Caption = dlgOpenFile.FileTitle + " Path: " + strOpen
End Sub

Private Sub Selectall_Click()
    R1.SelStart = 0
    R1.SelLength = Len(R1.Text)
    R1.SetFocus
End Sub

Private Sub Setlockatcurrentline_Click()
    Dim currLine As Long

    On Local Error Resume Next
    currLine = SendMessage(R1.hwnd, EM_LINEFROMCHAR, -1&, ByVal 0&) + 1
    'MsgBox Format$(currLine, "##,###")
 StatusBar1.Panels(1).Text = "Line: " + Format$(currLine, "##,###")
linelock = Format$(currLine, "##,###")

StatusBar1.Panels(5).Text = "Line locked: " + linelock
End Sub

Private Sub Splitx_Click()
Form2.Show


End Sub

Private Sub Start_Click()
xxx = "#!/usr/bin/perl" + vbCrLf
Clipboard.SetText xxx
R1.SelText = Clipboard.GetText
End Sub

Private Sub subedit_Click()

End Sub

Private Sub Timer1_Timer()
    Dim currLine As Long

    thechar = R1.SelStart
    Dim lineCount As Long
 StatusBar1.Panels(1).Text = "Char: " + thechar
  
    On Local Error Resume Next
    currLine = SendMessage(R1.hwnd, EM_LINEFROMCHAR, -1&, ByVal 0&) + 1
    lineCount = SendMessage(R1.hwnd, EM_GETLINECOUNT, 0&, ByVal 0&)
    'MsgBox Format$(currLine, "##,###")
 StatusBar1.Panels(2).Text = "Line: " + Format$(currLine, "##,###")
 StatusBar1.Panels(3).Text = "Total Lines: " + Format$(lineCount, "##,###")


End Sub



Private Sub Timer2_Timer()
selholder = R1.SelText
selmid = Mid(selholder, 1, 20)
Curseltext.Caption = "---> Current seltext: " + selmid + "..."
End Sub

Private Sub ToHTML_Click()
Dim holder As String
xxx = R1.SelText
xxx = Replace(xxx, "\" + Chr(34), Chr(34), 1, -1)
xxx = Replace(xxx, "\n", "", 1, -1)
xxx = Replace(xxx, "print " + Chr(34), "", 1, -1)
xxx = Replace(xxx, Chr(34) + ";", "", 1, -1)
R1.SelText = xxx
End Sub

Private Sub ToPerl_Click()
xxx = ""
Dim i As Variant
Dim MyString, myArray
MyString = Split(R1.SelText, vbCrLf, -1, 1)
For Each i In MyString
i = Replace(i, Chr(34), "\" + Chr(34))
If i = "" Or i = vbCrLf Then
Else
i = "print " + Chr(34) + i + "\n" + Chr(34) + ";"
End If
xxx = xxx + i + vbCrLf
Next
xxx = xxx + "/"
xxx = Replace(xxx, vbCrLf + "/", "")
R1.SelText = xxx
End Sub

Private Sub Undo_Click()
    'If gintIndex = 0 Then Exit Sub
    

    'gblnIgnoreChange = True
    'gintIndex = gintIndex - 1
    On Error Resume Next
    'R1.TextRTF = gstrStack(gintIndex)
    'gblnIgnoreChange = False
SendMessage R1.hwnd, EM_UNDO, gintIndex, 0
'SendMessage R1.hwnd, EM_SCROLLCARET, 0, 0
End Sub


Public Sub Update_Click()


Vars.List2.Clear
Vars.List1.Clear
If codecolen.Checked = False Then
Exit Sub
End If
R1.MousePointer = rtfArrowHourglass
Dim holdpos
'MsgBox sellenghtx
holdpos = R1.SelStart
If selstartx > 0 Then
R1.SelStart = selstartx
R1.SelLength = sellenghtx
Else
R1.SelStart = 0
R1.SelLength = Len(R1.Text)
End If
Dim alldata As String
alldata = R1.SelText
R1.Visible = False
Dim position
Dim startp
startp = 0
position = 1
Dim foundstr
Do While Not position = 0
position = InStr(startp + 1, alldata, "$", vbTextCompare)
If position > 0 Then
R1.SelStart = position - 1
R1.SelLength = 15
foundstr = R1.SelText
Dim currchar
Dim i As Integer
For i = 2 To 25
currchar = Mid(foundstr, i, 1)
If currchar = "{" Or _
currchar = " " Or _
currchar = ")" Or _
currchar = "," Or _
currchar = "/" Or _
currchar = "]" Or _
currchar = Chr(34) Or _
currchar = "'" Or _
currchar = "<" Or _
currchar = ";" Or _
currchar = "\" Or _
currchar = "-" Or _
currchar = "." Or _
currchar = ":" Or _
currchar = "(" Or _
currchar = "}" Or _
currchar = "[" Or _
currchar = "#" Then

Exit For
End If
Next
R1.SelLength = i - 1
R1.SelColor = &HC000&
Vars.List1.AddItem R1.SelText
startp = position
End If
Loop


startp = 0
position = 1
Do While Not position = 0
position = InStr(startp + 1, alldata, "print ", vbTextCompare)
If position > 0 Then
R1.SelStart = position - 1
R1.SelLength = 5
R1.SelColor = QBColor(9)
startp = position
End If
Loop
startp = 0
position = 1
Do While Not position = 0
position = InStr(startp + 1, alldata, "if ", vbTextCompare)
If position > 0 Then
R1.SelStart = position - 1
R1.SelLength = 2
R1.SelColor = QBColor(9)
startp = position
End If
Loop
startp = 0
position = 1
Do While Not position = 0
position = InStr(startp + 1, alldata, "else", vbTextCompare)
If position > 0 Then
R1.SelStart = position - 1
R1.SelLength = 4
R1.SelColor = QBColor(9)
startp = position
End If
Loop
startp = 0
position = 1
Do While Not position = 0
position = InStr(startp + 1, alldata, "open ", vbTextCompare)
If position > 0 Then
R1.SelStart = position - 1
R1.SelLength = 4
R1.SelColor = QBColor(9)
startp = position
End If
Loop
startp = 0
position = 1
Do While Not position = 0
position = InStr(startp + 1, alldata, "close", vbTextCompare)
If position > 0 Then
R1.SelStart = position - 1
R1.SelLength = 5
R1.SelColor = QBColor(9)
startp = position
End If
Loop
startp = 0
position = 1
Do While Not position = 0
position = InStr(startp + 1, alldata, "elsif ", vbTextCompare)
If position > 0 Then
R1.SelStart = position - 1
R1.SelLength = 4
R1.SelColor = QBColor(9)
startp = position
End If
Loop
R1.Visible = True
R1.SelStart = 0
R1.SelLength = 0
R1.SetFocus
R1.SelStart = holdpos
R1.SelColor = &H0&
R1.MousePointer = rtfDefault
selstartx = 0
sellenghtx = 0


'Dim i
Dim holder
holder = ""
'Vars.List1.Clear
If Vars.List1.ListCount = 0 Then
Exit Sub
Else
Vars.Show
End If

For i = 0 To Vars.List1.ListCount - 1
If Vars.List1.List(i) = holder Then
Else
Vars.List2.AddItem (Vars.List1.List(i))
End If
holder = Vars.List1.List(i)
Next

End Sub

Public Sub checkcolor()
Dim holdpos
holdpos = R1.SelStart
If thechar > 10 Then
R1.SelStart = thechar - 10
Else
R1.SelStart = 1
End If
R1.SelLength = thechar + 20
Dim alldata As String
alldata = R1.SelText
Dim varposit
varposit = InStr(1, alldata, "$", vbTextCompare)
''''''''''''''''''''
R1.SelStart = holdpos + varposit
R1.SelLength = 10
Dim foundstr
foundstr = R1.SelText

Dim currchar
Dim i As Integer
For i = 1 To 10
currchar = Mid(foundstr, i, 1)
If currchar = "{" Or _
currchar = " " Or _
currchar = ")" Or _
currchar = "," Or _
currchar = "/" Or _
currchar = "]" Or _
currchar = Chr(34) Or _
currchar = "'" Or _
currchar = "<" Or _
currchar = ";" Or _
currchar = "\" Or _
currchar = "-" Or _
currchar = "." Or _
currchar = ":" Or _
currchar = "(" Or _
currchar = "}" Or _
currchar = "[" Or _
currchar = "#" Then

Exit For
End If
Next
R1.SelStart = holdpos + varposit
R1.SelLength = i - 1
''''''''''''''''''''
R1.SelColor = &H80FF&
R1.SelStart = holdpos

End Sub

Public Sub selectword()
If thechar > 0 Then

End If
End Sub

Private Sub Variables_Click()
Vars.Show
End Sub
