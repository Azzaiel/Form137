Attribute VB_Name = "CommonHelper"
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
(ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, _
ByVal lpDirectory As String, _
ByVal nShowCmd As Long) As Long

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260

Private Declare Function SHBrowseForFolder Lib _
"shell32" (lpbi As BrowseInfo) As Long

Private Declare Function SHGetPathFromIDList Lib _
"shell32" (ByVal pidList As Long, ByVal lpBuffer _
As String) As Long

Private Declare Function lstrcat Lib "kernel32" _
Alias "lstrcatA" (ByVal lpString1 As String, ByVal _
lpString2 As String) As Long

Private Type BrowseInfo
   hWndOwner As Long
   pIDLRoot As Long
   pszDisplayName As Long
   lpszTitle As Long
   ulFlags As Long
   lpfnCallback As Long
   lParam As Long
   iImage As Long
End Type

Public Function selectDir(ownerHwnd As Long) As String
'Opens a Browse Folders Dialog Box that displays the
'directories in your computer
Dim lpIDList As Long ' Declare Varibles
Dim sBuffer As String
Dim szTitle As String
Dim tBrowseInfo As BrowseInfo

szTitle = "Hello World. Click on a directory and " & _
"it's path will be displayed in a message box"
' Text to appear in the the gray area under the title bar
' telling you what to do

With tBrowseInfo
   .hWndOwner = ownerHwnd ' Owner Form
   .lpszTitle = lstrcat(szTitle, "")
   .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
End With

lpIDList = SHBrowseForFolder(tBrowseInfo)

If (lpIDList) Then
   sBuffer = Space(MAX_PATH)
   SHGetPathFromIDList lpIDList, sBuffer
   sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
   selectDir = sBuffer
Else
  selectDir = ""
End If
End Function

Public Function openFile(FilePath As String, ownerHwnd As Long) As Boolean
     Dim dummy As Long
     
               'open the file using the default Editor or viewer.
     dummy = ShellExecute(ownerHwnd, "Open", FilePath & Chr$(0), Chr$(0), _
                                          Left$(FilePath, InStr(FilePath, "\")), 0)
     openFile = dummy
     
End Function

Public Function extractStringValue(value As Object) As String
  If (Not IsNull(value)) Then
    extractStringValue = value
  Else
    extractStringValue = ""
  End If
End Function
Public Function isFunctionAscii(ascii As Integer) As Boolean
  If (ascii = 13 Or ascii = 8 Or ascii = 32) Then
    isFunctionAscii = True
  Else
    isFunctionAscii = False
  End If
End Function
Public Function isNumberAscii(ascii As Integer) As Boolean
  If (ascii >= 48 And ascii <= 57) Then
    isNumberAscii = True
  Else
    isNumberAscii = False
  End If
End Function
Public Function extractDateValue(value As Object) As String
  If (Not IsNull(value)) Then
    extractDateValue = Format(value, Constants.DEFAULT_FORMAT)
  Else
    extractDateValue = ""
  End If
End Function
Public Function hasValidValue(value As String) As Boolean
   Dim isValid As Boolean
   isValid = True
   If (Not IsNull(value)) Then
   
     If (IsNumeric(value)) Then
       isValid = val(value) > 0
     Else
       isValid = Trim(value) <> vbNullString
     End If
   End If
   hasValidValue = isValid
End Function
Public Sub sendWarning(txtBox As TextBox, errMsg As String)
  MsgBox errMsg, vbCritical
  txtBox.BackColor = vbRed
  txtBox.ForeColor = vbWhite
  txtBox.SetFocus
End Sub
Public Sub sendComboBoxWarning(cmbBox As ComboBox, errMsg As String)
  MsgBox errMsg, vbCritical
  cmbBox.BackColor = vbRed
  cmbBox.ForeColor = vbWhite
End Sub
Public Sub toDefaultSkin(txtBox As TextBox)
  txtBox.BackColor = vbWhite
  txtBox.ForeColor = vbBlack
End Sub
Public Sub toComboBoxDefaultSkin(cmbBox As ComboBox)
  cmbBox.BackColor = vbWhite
  cmbBox.ForeColor = vbBlack
End Sub
Public Function getFileName(flname As String) As String

    Dim posn As Integer, i As Integer
    Dim fName As String

    posn = 0
    For i = 1 To Len(flname)
        If (Mid(flname, i, 1) = "\") Then posn = i
    Next i

    fName = Right(flname, Len(flname) - posn)

    getFileName = fName
    
End Function
Public Function getImgPath() As String
  getImgPath = App.Path & "\img"
End Function

Public Function getTemplatesPath() As String
  getTemplatesPath = App.Path & "\template"
End Function
Public Function getTempPath() As String
  getTempPath = App.Path & "\tmp"
End Function


