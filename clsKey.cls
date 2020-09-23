VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VBKeyWord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Keyword"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Option Compare Text
Dim KeyWord As New Collection
Private m_Delim As String
Const DefaultDelim As String = ","
Public Event Erreur(ByVal FunctionName As String, ByVal Number As Long, ByVal Description As String, ByVal DataError As String)
Private Sub Class_Initialize()
    m_Delim = DefaultDelim
    KeyWord.Add "If", "If"
    KeyWord.Add "Endif", "Endif"
    KeyWord.Add "Print", "Print"
    KeyWord.Add "Debug", "Debug"
    KeyWord.Add "Sub", "Sub"
    KeyWord.Add "Function", "Function"
    KeyWord.Add "End", "End"
    KeyWord.Add "Public", "Public"
    KeyWord.Add "Nothing", "Nothing"
    KeyWord.Add "Dim", "Dim"
    KeyWord.Add "String", "String"
    KeyWord.Add "Long", "Long"
    KeyWord.Add "Option", "Option"
    KeyWord.Add "Compare", "Compare"
    KeyWord.Add "On", "On"
    KeyWord.Add "Error", "Error"
    KeyWord.Add "Set", "Set"
    KeyWord.Add "Byval", "Byval"
    KeyWord.Add "ByRef", "Byref"
    KeyWord.Add "Optional", "Optional"
    KeyWord.Add "Variant", "Variant"
End Sub
Private Sub Class_Terminate()
    Set KeyWord = Nothing
End Sub
Public Sub Add(Item As Variant, Optional ByVal Key As Variant, Optional ByVal Before As Variant, Optional ByVal After As Variant)
    On Error GoTo err_Occur
    KeyWord.Add Item, Key, Before, After
    On Error GoTo 0
err_Continu:
    Exit Sub
err_Occur:
    RaiseEvent Erreur("Add", Err.Number, Err.Description, "")
    Resume err_Continu
End Sub
Public Sub RemoveKey(ByVal Key As String)
    On Error GoTo err_Occur
    KeyWord.Remove Key
    On Error GoTo 0
err_Continu:
    Exit Sub
err_Occur:
    RaiseEvent Erreur("RemoveKey", Err.Number, Err.Description, Key)
    Resume err_Continu
End Sub




Public Sub Remove(ByVal IndexOrKey As Variant)


    On Error GoTo err_Occur
    KeyWord.Remove IndexOrKey
    On Error GoTo 0
err_Continu:
    Exit Sub
err_Occur:
    RaiseEvent Erreur("Remove", Err.Number, Err.Description, IndexOrKey)
    Resume err_Continu
End Sub




Public Sub RemoveIndex(ByVal Index As Long)


    On Error GoTo err_Occur


    If Index <= KeyWord.Count Then
        KeyWord.Remove Index
    Else
        RaiseEvent Erreur("RemoveIndex", 9, "Subscript out of range, Max=" + CStr(KeyWord.Count), Index)
    End If


    On Error GoTo 0
err_Continu:
    Exit Sub
err_Occur:
    MsgBox Err.Number
    RaiseEvent Erreur("RemoveIndex", Err.Number, Err.Description, Index)
    Resume err_Continu
End Sub




Public Sub RemoveAll()


    If KeyWord.Count = 0 Then Exit Sub
    Dim element As Variant


    For Each element In KeyWord
        KeyWord.Remove 1
    Next element


End Sub




Public Property Get Count() As Long


    On Error GoTo err_Occur
    Count = KeyWord.Count
    On Error GoTo 0
err_Continu:
    Exit Function
err_Occur:
    RaiseEvent Erreur("Count", Err.Number, Err.Description, "")
    Resume err_Continu
End Property




Public Function Item(ByVal IndexOrKey As Variant) As Variant
Attribute Item.VB_UserMemId = 0
    On Error GoTo err_Occur
    Item = KeyWord.Item(IndexOrKey)
    On Error GoTo 0
err_Continu:
    Exit Function
err_Occur:
    RaiseEvent Erreur("Item", Err.Number, Err.Description, IndexOrKey)
    Resume err_Continu
End Function




Public Function IfItemIsThere(ByVal Index As Long) As Boolean


    Dim temp As Variant
    On Error GoTo err_Occur
    temp = KeyWord.Item(Index)
    On Error GoTo 0
    IfItemIsThere = True
err_Continu:
    Exit Function
err_Occur:
    IfItemIsThere = False
    Resume err_Continu
End Function
Public Function IfKeyIsThere(ByVal Key As String) As Boolean


    Dim temp As Variant
    On Error GoTo err_Occur
    temp = KeyWord.Item(Key)
    On Error GoTo 0
    IfKeyIsThere = True
err_Continu:
    Exit Function
err_Occur:
    IfKeyIsThere = False
    Resume err_Continu
End Function
Public Property Get DelimStringDataError() As String
    DelimStringDataError = m_Delim
End Property
Public Property Let DelimStringDataError(ByVal NewDelim As String)
    m_Delim = NewDelim
End Property

' NewEnum must return the IUnknown interface of a' collection's enumerator.
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = KeyWord.[_NewEnum]
End Function

