Attribute VB_Name = "Module1"
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Declare Function SendMessage Lib _
    "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, _
     ByVal wMsg As Long, _
     ByVal wParam As Long, _
     lParam As Any) As Long

Public active As Boolean
Public xcoord
Public ycoord
Public starttext(50)
Public perlpath
Public Const EM_LINEFROMCHAR = &HC9
Public Const EM_GETLINECOUNT = &HBA
Public Const EM_SCROLLCARET = &HB7
Public Const EM_UNDO = &HC7
Public Const EM_GETLINE = &HC4





Public Sub colorword(wordtocolor As String, wordnr As Integer, farge As String, xkey As Integer)

Dim holder
Dim wordl
wordl = Len(wordtocolor) - 1
holder = starttext(wordnr)

starttext(wordnr) = Mid(starttext(wordnr), 2, wordl) + Chr(xkey)
Form1.Caption = wordtocolor
If holder = wordtocolor Then
Form1.R1.SelStart = Form1.R1.SelStart - Len(wordtocolor)
Form1.R1.SelLength = 6
Form1.R1.SelColor = farge
Form1.R1.SelText = wordtocolor
Form1.R1.SelColor = &H0&
End If
End Sub

Public Sub colorit(cchar As String)

Form1.R1.Find "open", cchar - 5, cchar + 5
Form1.R1.SelColor = &H8000000F
Form1.R1.SelStart = cchar
End Sub





