VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form adviserAddStudent 
   BackColor       =   &H8000000E&
   Caption         =   "Form1"
   ClientHeight    =   7395
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13650
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   ScaleHeight     =   7395
   ScaleWidth      =   13650
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "CLOSE"
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
      Left            =   6120
      TabIndex        =   13
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   " Add All >>"
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
      Left            =   6120
      TabIndex        =   10
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "<< Remove All"
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
      Left            =   6120
      TabIndex        =   9
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<< Remove"
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
      Left            =   6120
      TabIndex        =   8
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add >>"
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
      Left            =   6120
      TabIndex        =   7
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00808080&
      Height          =   735
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   15375
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Level:"
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
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Section:"
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
         Left            =   4320
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lbl_level 
         BackStyle       =   0  'Transparent
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
         Left            =   1080
         TabIndex        =   4
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label lbl_section 
         BackStyle       =   0  'Transparent
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
         Left            =   5400
         TabIndex        =   3
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000E&
      Caption         =   "Students In Section"
      Height          =   6495
      Left            =   7440
      TabIndex        =   1
      Top             =   720
      Width           =   5895
      Begin MSDataGridLib.DataGrid dg_current_stud 
         Height          =   6135
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   10821
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   19
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   5895
      Begin MSDataGridLib.DataGrid dg_available_stud 
         Height          =   5055
         Left            =   120
         TabIndex        =   11
         Top             =   1320
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   8916
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   19
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "adviserAddStudent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rs_current_stud As New ADODB.Recordset
Private rs_available_stud As New ADODB.Recordset
Private rs_tmp As New ADODB.Recordset
Private sql_query As String
Private Sub populateCurrentStudent()
  sql_query = "Select a.STUDENT_ID as LRN, CONCAT(a.LAST_NAME, ', ' , a.FIRST_NAME) as Name, a.GENDER " & _
              "From tbl_student a, tbl_student_level b " & _
              "Where b.ID = a.STUDENT_ID " & _
              "      And b.SY= '" & mainteacherform.cmb_sy.Text & "' " & _
              "      And b.LVL_NAME = '" & lbl_level & "' " & _
              "      And b.SECTION_NAME = '" & lbl_section & "' " & _
              "Order By a.Gender "
  Call set_datagrid(dg_current_stud, rs_current_stud, sql_query)
  With dg_current_stud
    .Columns(0).Width = 1550
    .Columns(2).Width = 1000
  End With
End Sub
Private Sub Command1_Click()
  If (rs_available_stud.RecordCount > 0) Then
    sql_query = "Select * from tbl_student_level where 1 = 2"
    Call mysql_select(rs_tmp, sql_query)
    rs_tmp.AddNew
    rs_tmp!id = rs_available_stud!LRN
    rs_tmp!lvl_name = lbl_level
    rs_tmp!section_name = lbl_section
    rs_tmp!status = "ENROLLED"
    rs_tmp!SY = mainteacherform.cmb_sy.Text
    rs_tmp.Update
    MsgBox "Student added to Section!", vbInformation
    Call Form_Load
  End If
End Sub

Private Sub Command2_Click()
  If (rs_current_stud.RecordCount > 0) Then
    sql_query = "Select * from tbl_student_level " & _
                "Where ID = '" & rs_current_stud!LRN & "' " & _
                "      And SY = '" & mainteacherform.cmb_sy.Text & "'"
    Call mysql_select(rs_tmp, sql_query)
    If (rs_tmp.RecordCount > 0) Then
      rs_tmp.Delete
      MsgBox "Student Removed from Section", vbInformation
    End If
    Call Form_Load
  End If
End Sub
End Sub

Private Sub Command5_Click()
  Unload Me
  Call masterlistadvisoriesform.Form_Load
End Sub

Public Sub Form_Load()
   If (lbl_level <> vbNullString) Then
     Call populateCurrentStudent
     Call populateAvailableStundet
   End If
End Sub
Private Sub populateAvailableStundet()
  sql_query = "Select a.STUDENT_ID as LRN, CONCAT(a.LAST_NAME, ', ' , a.FIRST_NAME) as Name, a.GENDER " & _
              "From tbl_student a " & _
              "Where a.STUDENT_ID not in (Select ID from tbl_student_level) " & _
              "Order By a.Gender"
  Call set_datagrid(dg_available_stud, rs_available_stud, sql_query)
  With dg_available_stud
    .Columns(0).Width = 1550
    .Columns(2).Width = 1000
  End With
End Sub