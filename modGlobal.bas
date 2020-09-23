Attribute VB_Name = "modGlobal"
Option Explicit
Option Compare Text
Public Declare Function SendMessage Lib _
    "user32" Alias "SendMessageA" _
    (ByVal hWnd As Long, _
     ByVal wMsg As Long, _
     ByVal wParam As Long, _
     lParam As Any) As Long

Public Const LB_FINDSTRING = &H18F
Public Const LB_FINDSTRINGEXACT = &H1A2
Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Type POINTAPI
        x As Long
        y As Long
End Type
Public Declare Function SetWindowPos Lib "user32" _
    (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, _
    ByVal x As Long, ByVal y As Long, ByVal cx As Long, _
    ByVal cy As Long, ByVal wFlags As Long) As Long

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE


Public Function FindKey(ctl As ListBox, ByVal strKey As String) As OLE_COLOR
    ctl.ListIndex = SendMessage(ctl.hWnd, LB_FINDSTRINGEXACT, 0&, ByVal (strKey))
    If Not ctl.ListIndex = -1 Then
        
        FindKey = ctl.ItemData(ctl.ListIndex)
    Else
        FindKey = 0
    End If
End Function

Public Function InstrBak(lngStartPos As Long, strSource As String, strSearch As String, Optional intCompareMode As Integer, Optional ByVal lngSearchCount As Long) As Long

  '*******************************************************************************
  '
  ' DESCRIPTION
  '     Same as Instr but starts the search in reverse.
  '
  ' ARGUMENTS
  '     lngStartPos    = Position to start the search from
  '                      (-1 starts from the very end)
  '     strSource      = Source string
  '     strSearch      = Search string
  '     intCompareMode = Optional. vbTextCompare, vbBinaryCompare,
  '                      vbDatabaseCompare. If missing, defaults to vbBinaryCompare
  '     lngSearchCount = Optional.  Allows you to define which strSearch to
  '                      return.  For example, if you are reverse searching the
  '                      "\" character but what the position of the 2 one found,
  '                      you would set it to 2.  If not specifed, the position of
  '                      the first occurrence is returned.
  '
  ' RETURNS
  '     Position of search value.
  '
  ' DEPENDENCIES
  '     None
  '
  ' REMARKS
  '
  '*******************************************************************************
  
  On Local Error Resume Next
  
  Dim lngFound As Long
  Dim lngCurPos As Long
  Dim lngSourceLen As Long
  Dim lngN As Long
  Dim intCounter As Integer
  Dim lngPrevI As Long
         
  lngSourceLen = Len(strSource)
  
  If lngStartPos = -1 Then
    lngCurPos = lngSourceLen
  Else
    lngCurPos = lngStartPos
  End If
  
  For lngN = lngCurPos To 1 Step -1
    lngFound = InStr(lngN, strSource, strSearch, intCompareMode)
    If lngFound Then
      If lngSearchCount = 0 Then
        Exit For
      ElseIf lngFound <> lngPrevI Then
        intCounter = intCounter + 1
        If lngSearchCount = intCounter Then
          Exit For
        Else
          lngPrevI = lngFound
        End If
      End If
    End If
  Next
       
  InstrBak = lngFound
       
End Function
Public Sub SetTopMost(ByVal lHwnd As Long, ByVal bTopMost As Boolean)
'
' Set the hwnd of the window topmost or not topmost
'
    Dim lUseVal  As Long
    Dim lRet As Long
    
    lUseVal = IIf(bTopMost, HWND_TOPMOST, HWND_NOTOPMOST)
    
    lRet = SetWindowPos(lHwnd, lUseVal, 0, 0, 0, 0, FLAGS)
    
    If lRet < 0 Then
'
' Couldn't do operation - handle error here
'
'        DisplayWinAPIError lRet
    End If

End Sub

