VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form masterlistadvisoriesform 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Masterlist"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "masterlistadvisoriesform.frx":0000
   ScaleHeight     =   6015
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmb_category 
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
      ItemData        =   "masterlistadvisoriesform.frx":AFCC2
      Left            =   3360
      List            =   "masterlistadvisoriesform.frx":AFCCF
      TabIndex        =   9
      Top             =   720
      Width           =   3135
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   -120
      Width           =   8775
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
         TabIndex        =   6
         Top             =   240
         Width           =   3135
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
         TabIndex        =   5
         Top             =   240
         Width           =   2895
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
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
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
         TabIndex        =   3
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.CommandButton cmd_print 
      Height          =   615
      Left            =   3600
      Picture         =   "masterlistadvisoriesform.frx":AFCEF
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5400
      Width           =   1575
   End
   Begin MSDataGridLib.DataGrid dg_masterlist 
      Height          =   4095
      Left            =   240
      TabIndex        =   7
      Top             =   1200
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   7223
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
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Add/Remove Student"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   5520
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Sort by:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   8
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label lbl_set_grade 
      BackStyle       =   0  'Transparent
      Caption         =   "Encode character grade."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   6120
      TabIndex        =   1
      Top             =   5520
      Visible         =   0   'False
      Width           =   2655
   End
End
Attribute VB_Name = "masterlistadvisoriesform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rs_masterlist As New ADODB.Recordset
Dim rs_grade As New ADODB.Recordset
Public sel_student_name As String
Public sel_lrn As String


Private Sub cmb_category_Click()
    Dim col_order As String
    Select Case (cmb_category.Text)
        Case "LRN"
            col_order = "a.student_id"
        Case "Last Name"
            col_order = "a.last_name"
        Case "First Name"
            col_order = "a.first_name"
    End Select
    
    Call set_datagrid(dg_masterlist, rs_masterlist, _
                                            "SELECT @index := @index + 1 as No," _
                                                & "masterlist.* " _
                                            & "FROM " _
                                                & "(SELECT " _
                                                    & "a.student_id as LRN, " _
                                                    & "a.last_name as Last_Name, a.First_Name,a.Middle_Name " _
                                                & "FROM " _
                                                    & "tbl_student a " _
                                                & "LEFT JOIN " _
                                                    & "tbl_student_level b " _
                                                & "ON " _
                                                    & "a.student_id = b.ID and b.sy = '" & mainteacherform.cmb_sy.Text & "' " _
                                                & "WHERE " _
                                                    & " b.section_name = '" & myadvisoriesform.rs_advisories.Fields("Section") & "' ORDER BY " & col_order & " ASC) masterlist" _
                                            & " JOIN " _
                                                & "(SELECT @index :=0) c ")
    dg_masterlist.Columns(0).Width = 400
End Sub

Private Sub cmb_category_KeyUp(KeyCode As Integer, Shift As Integer)
    MsgBox "Please select category from the list."
    cmb_category.Text = ""
End Sub

Private Sub cmd_print_Click()
    If rs_masterlist.RecordCount = 0 Then
        MsgBox "No record."
        Exit Sub
    End If
      'dr_advisories.Sections(2).Controls("lbl_sy").Caption = mainteacherform.lbl_sy.Caption
      dr_advisories.Sections(2).Controls("lbl_date").Caption = Format(Now, "mmmm dd, yyyy") & "/ " & Time
        dr_advisories.Sections(2).Controls("lbl_no").Caption = rs_masterlist.RecordCount
        dr_advisories.Sections(2).Controls("lbl_section").Caption = lbl_level.Caption & " - " & lbl_section.Caption
         Set dr_advisories.DataSource = rs_masterlist
        dr_advisories.Show vbModal, Me
End Sub
Private Sub encodeStudentCharacterGrade()
  
End Sub

Private Sub dg_masterlist_DblClick()
   sel_lrn = rs_masterlist!LRN
   sel_student_name = rs_masterlist!LAST_NAME & ", " & rs_masterlist!FIRST_NAME
   Call load_form(CharEncodePeriodSelect, True)
End Sub

Public Sub Form_Load()
     Call set_datagrid(dg_masterlist, rs_masterlist, _
                                        "SELECT @index := @index + 1 as No," _
                                            & "a.student_id as LRN, a.last_name as Last_Name, a.First_Name,a.Middle_Name FROM tbl_student a LEFT JOIN tbl_student_level b ON a.student_id = b.ID and b.sy = '" & mainteacherform.cmb_sy.Text & "' JOIN(SELECT @index :=0) c WHERE  b.section_name = '" & myadvisoriesform.rs_advisories.Fields("Section") & "'")
     dg_masterlist.Columns(0).Width = 400
End Sub
Private Sub Label4_Click()
  adviserAddStudent.lbl_level = lbl_level
  adviserAddStudent.lbl_section = lbl_section
  Call adviserAddStudent.Form_Load
  Call load_form(adviserAddStudent, True)
End Sub

Private Sub lbl_set_grade_Click()
      If rs_masterlist.RecordCount = 0 Then
        MsgBox "No record."
        Exit Sub
    End If
    characterencodeform.lbl_level.Caption = lbl_level.Caption
    characterencodeform.lbl_section.Caption = lbl_section.Caption
      Call set_datagrid(characterencodeform.dg_grades, rs_grade, _
                                        "SELECT " _
                                            & "a.student_id as LRN, a.last_name as Last_Name, a.First_Name FROM tbl_student a LEFT JOIN tbl_student_level b ON a.student_id = b.ID WHERE b.section_name = '" & section & "'")
         
    Call load_form(characterencodeform, True)
End Sub
