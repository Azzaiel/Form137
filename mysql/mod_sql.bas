Attribute VB_Name = "mod_sql"

Option Explicit

Private Const SERVER = "localhost"
Private Const USERNAME = "root"
Private Const PASSWORD = "root"
Private Const DATABASE = "db_form137"
Public Declare Function GetShortPathName Lib "kernel32" _
    Alias "GetShortPathNameA" _
    (ByVal lpszLongPath As String, _
    ByVal lpszShortPath As String, _
    ByVal cchBuffer As Long) As Long

Public Function GetShortName(sFile As String) As String
    Dim sShortFile As String * 67
    Dim lResult As Long

    'Make a call to the GetShortPathName API
    lResult = GetShortPathName(sFile, sShortFile, _
    Len(sShortFile))

    'Trim out unused characters from the string.
    GetShortName = Left$(sShortFile, lResult)

End Function

Public Sub backup_db(my_path)
        Shell "cmd.exe /c """ & GetShortName(App.Path) & "\mysql\mysqldump.exe"" -h" & SERVER & " -p" & PASSWORD & " -u" & USERNAME & " " & DATABASE & " > " & my_path & ""
End Sub
Public Sub restore_db(my_path)
        Shell "cmd.exe /c """ & GetShortName(App.Path) & "\mysql\mysql.exe"" -h" & SERVER & " -p" & PASSWORD & " -u" & USERNAME & " " & DATABASE & " < " & my_path & ""
End Sub

