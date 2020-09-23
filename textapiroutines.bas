Attribute VB_Name = "TextAPIs"
Option Explicit

Public Const EM_LINEFROMCHAR = &HC9
Public Const EM_LINEINDEX = &HBB
Public Const EM_LINELENGTH = &HC1
Public Const EM_GETLINE = &HC4
Public Const EM_GETLINECOUNT = &HBA
Public Const EM_GETFIRSTVISIBLELINE = &HCE
Public Const EM_GETRECT = &HB2
Public Const WM_GETFONT = &H31
Public Const WM_VSCROLL = &H115
Public Const EM_CHARFROMPOS = &HD7
Public Const SB_LINEUP = 0
Public Const SB_LINEDOWN = 1
Public Const EM_CANUNDO = &HC6
Private Const EM_UNDO = &HC7


Public Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Public Type TEXTMETRIC
  tmHeight As Long
  tmAscent As Long
  tmDescent As Long
  tmInternalLeading As Long
  tmExternalLeading As Long
  tmAveCharWidth As Long
  tmMaxCharWidth As Long
  tmWeight As Long
  tmOverhang As Long
  tmDigitizedAspectX As Long
  tmDigitizedAspectY As Long
  tmFirstChar As Byte
  tmLastChar As Byte
  tmDefaultChar As Byte
  tmBreakChar As Byte
  tmItalic As Byte
  tmUnderlined As Byte
  tmStruckOut As Byte
  tmPitchAndFamily As Byte
  tmCharSet As Byte
End Type

Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageByNum Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As String) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hdc As Long, lpMetrics As TEXTMETRIC) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long

Public Sub GetLineCol(ByVal hWnd As Long, ByVal nAbsValue As Long, ByRef nLine As Long, ByRef nCol As Long)
Dim lLine As Long
Dim lIndex As Long
  'Get the new line number
  lLine = SendMessage(hWnd, EM_LINEFROMCHAR, nAbsValue, 0)
  lIndex = SendMessage(hWnd, EM_LINEINDEX, lLine, 0)
  nLine = lLine
  nCol = nAbsValue - lIndex
End Sub

Public Function GetFirstVisibleLine(ByVal hWnd As Long) As Long
  GetFirstVisibleLine = SendMessage(hWnd, EM_GETFIRSTVISIBLELINE, 0, 0&)
End Function

Public Function GetLastVisibleLine(ByVal hWnd As Long) As Long
  GetLastVisibleLine = GetVisibleLines(hWnd) + GetFirstVisibleLine(hWnd) - 1
End Function

Public Function GetVisibleLines(ByVal hWnd As Long) As Long
Dim rc As RECT
Dim hdc As Long
Dim lFont As Long
Dim OldFont As Long
Dim di As Long
Dim tm As TEXTMETRIC
Dim lc As Long
  lc = SendMessage(hWnd, EM_GETRECT, 0, rc)
  lFont = SendMessage(hWnd, WM_GETFONT, 0, 0)
  hdc = GetDC(hWnd)
  If lFont <> 0 Then OldFont = SelectObject(hdc, lFont)
  di = GetTextMetrics(hdc, tm)
  If lFont <> 0 Then lFont = SelectObject(hdc, OldFont)
  GetVisibleLines = (rc.Bottom - rc.Top) / tm.tmHeight
  di = ReleaseDC(hWnd, hdc)
End Function

Public Function GetLineText(EditControl As Control, ByVal nLineNumber As Long) As String
Dim lIndex As Long
Dim lc As Long
Dim LineBuffer As String

  'Ensure Linenumber valid
  If nLineNumber < 0 Then nLineNumber = 0
  lc = GetMaxLines(EditControl.hWnd)
  If nLineNumber > lc - 1 Then nLineNumber = lc - 1
  'Return starting char pos on passed line number
  lIndex = SendMessageByNum(EditControl.hWnd, EM_LINEINDEX, nLineNumber, 0&)
  'Return the length of the passed line number
  lc = SendMessageByNum(EditControl.hWnd, EM_LINELENGTH, lIndex, 0&) + 1
  'Extract the text contained on the passed line number
  LineBuffer = Mid$(EditControl.Text, lIndex + 1, lc)
  GetLineText = Trim$(LineBuffer)

End Function

Public Sub SelectLineText(EditControl As Control, ByVal nLineNumber As Long)
Dim lIndex As Long
Dim lc As Long
Dim LineBuffer As String
    
  'Ensure Linenumber valid
  If nLineNumber < 0 Then nLineNumber = 0
  lc = GetMaxLines(EditControl.hWnd)
  If nLineNumber > lc - 1 Then nLineNumber = lc - 1
  'Return starting char pos on passed line number
  lIndex = SendMessageByNum(EditControl.hWnd, EM_LINEINDEX, nLineNumber, 0&)
  'Return the length of the passed line number
  lc = SendMessageByNum(EditControl.hWnd, EM_LINELENGTH, lIndex, 0&) + 1
  EditControl.SelStart = lIndex
  EditControl.SelLength = lc
  Debug.Print EditControl.SelText
End Sub

Public Function GetMaxLines(ByVal hWnd As Long) As Long
  GetMaxLines = SendMessage(hWnd, EM_GETLINECOUNT, 0, 0&)
End Function
Public Function GetTotalSelectedLine(ByVal hWnd As Long, ByVal strSel As String) As Long
    GetTotalSelectedLine = SendMessage(hWnd, EM_GETLINECOUNT, Len(strSel), 0)
End Function

Public Function CanUndo(ByVal hWnd As Long) As Long
    CanUndo = SendMessage(hWnd, EM_CANUNDO, 0, 0)
End Function
Public Sub Undo(ByVal hWnd As Long)
    SendMessage hWnd, EM_UNDO, 0, 0
End Sub
Public Sub Redo(ByVal hWnd As Long)
    SendMessage hWnd, EM_redo, 0, 0
End Sub
