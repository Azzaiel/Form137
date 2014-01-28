VERSION 5.00
Begin VB.Form adviserSelectSubject 
   ClientHeight    =   1665
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   5235
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1665
   ScaleWidth      =   5235
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5295
      Begin VB.ComboBox cmb_subject 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "adviserSelectSubject.frx":0000
         Left            =   2160
         List            =   "adviserSelectSubject.frx":0002
         TabIndex        =   3
         Top             =   480
         Width           =   2535
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Submit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1320
         TabIndex        =   2
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3120
         TabIndex        =   1
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Select Subject:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   4
         Top             =   480
         Width           =   1695
      End
   End
End
Attribute VB_Name = "adviserSelectSubject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private subj_code_list As Variant
Private Sub cmb_subject_KeyPress(KeyAscii As Integer)
  If (KeyAscii = 13) Then
    Call Command1_Click
  End If
End Sub
Private Sub Command1_Click()
  If (cmb_subject.Text <> vbNullString) Then
    bulkSubjEncode.lbl_level = masterlistadvisoriesform.lbl_level
    bulkSubjEncode.lbl_section = masterlistadvisoriesform.lbl_section
    bulkSubjEncode.lbl_subject = cmb_subject.Text
    bulkSubjEncode.subj_code = subj_code_list(cmb_subject.ListIndex)
    Dim sql_query As String
    sql_query = ""
    Call load_form(bulkSubjEncode, True)
  Else
    MsgBox "Please select a Subject", vbCritical
  End If
End Sub

Private Sub Command2_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  Dim sql_query As String
  sql_query = "Select Subject_Name, Subject_Code " & _
              "From tbl_subjectset " & _
              "Where lvl_name = '" & masterlistadvisoriesform.lbl_level & "' " & _
              "      and section_name = '" & masterlistadvisoriesform.lbl_section & "' " & _
              "Order By Subject_Name "

  Call mysql_select(public_rs, sql_query)

  adviserSelectSubject.cmb_subject.Clear
  ReDim subj_code_list(0 To public_rs.RecordCount) As String
  Dim index As Integer
  index = 0
  While Not public_rs.EOF
    adviserSelectSubject.cmb_subject.AddItem public_rs!Subject_Name
    subj_code_list(index) = public_rs!Subject_Code
    index = index + 1
    public_rs.MoveNext
  Wend
 
End Sub

