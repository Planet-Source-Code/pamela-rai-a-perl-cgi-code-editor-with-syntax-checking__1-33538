Attribute VB_Name = "Module2"
Private Declare Function GetShortPathName Lib "KERNEL32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
    
Public Function GetShortFileName(ByVal FileName As String) As String
    'converts a long file and path name to o
    '     ld DOS format
    'PARAMETERS
    'FileName = the path or filename to conv
    '     ert
    'RETURNS
    'String = the DOS compatible name for th
    '     at particular FileName
    'USES
    'KERNEL32 API call GetShortPathNameA
    'CONSIDERATIONS
    'short filename equivalents should only
    '     be used with non
    'Win95 programs
    Dim rc As Long
    Dim ShortPath As String
    Const PATH_LEN& = 164
    'get the short filename
    ShortPath = String$(PATH_LEN + 1, 0)
    rc = GetShortPathName(FileName, ShortPath, PATH_LEN)
    GetShortFileName = Left$(ShortPath, rc)
End Function

