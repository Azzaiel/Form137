VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form gradeencodeform 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Encode Grades"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "gradeencodeform.frx":0000
   ScaleHeight     =   5985
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   -120
      Width           =   8775
      Begin VB.Label lbl_subject_name 
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
         Left            =   5280
         TabIndex        =   9
         Top             =   600
         Width           =   3375
      End
      Begin VB.Label lbl_subject_code 
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
         TabIndex        =   8
         Top             =   600
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
         TabIndex        =   7
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
         TabIndex        =   6
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
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
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Code:"
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
         TabIndex        =   4
         Top             =   600
         Width           =   975
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
         TabIndex        =   3
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
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.CommandButton cmd_print 
      Height          =   615
      Left            =   3720
      Picture         =   "gradeencodeform.frx":1B3CE
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5280
      Width           =   1575
   End
   Begin MSDataGridLib.DataGrid dg_grades 
      Height          =   3855
      Left            =   240
      TabIndex        =   10
      Top             =   1320
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   6800
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
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Double-click to encode the grade,"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   960
      Width           =   3735
   End
End
Attribute VB_Name = "gradeencodeform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rs_grade As New ADODB.Recordset
Public rs_grade2 As New ADODB.Recordset
Public rs_1st As New ADODB.Recordset
Public rs_2nd As New ADODB.Recordset
Public rs_3rd As New ADODB.Recordset
Public rs_4th As New ADODB.Recordset
Public rs_final As New ADODB.Recordset
Dim sql_string As String



Private Sub cmd_print_Click()
    If rs_grade.RecordCount = 0 Then
        MsgBox "No record."
        Exit Sub
    End If
       dr_masterlist.Sections(2).Controls("lbl_sy").Caption = mainteacherform.lbl_sy.Caption
        dr_masterlist.Sections(2).Controls("lbl_level").Caption = lbl_level.Caption
        dr_masterlist.Sections(2).Controls("lbl_section").Caption = lbl_section.Caption
        dr_masterlist.Sections(2).Controls("lbl_subject").Caption = lbl_subject_name.Caption
         Set dr_masterlist.DataSource = rs_grade
        dr_masterlist.Show vbModal, Me
End Sub

Private Sub dg_grades_DblClick()
If dg_grades Is Nothing Then
    MsgBox "No record."
    Exit Sub
End If
     gradeform.lbl_id.Caption = rs_grade.Fields("LRN")
     gradeform.lbl_name.Caption = rs_grade.Fields("First_Name").Value & " " & rs_grade.Fields("Last_Name").Value
     gradeform.lbl_section2.Caption = lbl_section.Caption
     gradeform.lbl_subject_name2.Caption = lbl_subject_code.Caption
       Call mysql_select(public_rs, "SELECT * FROM tbl_grade WHERE ID='" & rs_grade.Fields("LRN").Value & "' AND subject_code = '" & lbl_subject_code.Caption & "'")
    If public_rs.RecordCount = 0 Then
        gradeform.txt_1st.Text = "0"
        gradeform.txt_2nd.Text = "0"
        gradeform.txt_3rd.Text = "0"
        gradeform.txt_4th.Text = "0"
        gradeform.lbl_final.Text = "No grade"
        gradeform.lbl_remark_1st.Text = "No grade"
        gradeform.lbl_remark_2nd.Text = "No grade"
        gradeform.lbl_remark_3rd.Text = "No grade"
        gradeform.lbl_remark_4th.Text = "No grade"
        gradeform.lbl_remark_final.Text = "No grade"
    Else
        Call mysql_select(rs_1st, "SELECT * FROM tbl_grade WHERE ID='" & rs_grade.Fields("LRN").Value & "' AND subject_code = '" & lbl_subject_code.Caption & "'AND Period='1st Grading'")
        gradeform.txt_1st.Text = rs_1st.Fields("grade").Value
        gradeform.lbl_remark_1st.Text = rs_1st.Fields("remark").Value
        Call mysql_select(rs_2nd, "SELECT * FROM tbl_grade WHERE ID='" & rs_grade.Fields("LRN").Value & "' AND subject_code = '" & lbl_subject_code.Caption & "'AND Period='2nd Grading'")
        gradeform.txt_2nd.Text = rs_2nd.Fields("grade").Value
        gradeform.lbl_remark_2nd.Text = rs_2nd.Fields("remark").Value
        Call mysql_select(rs_3rd, "SELECT * FROM tbl_grade WHERE ID='" & rs_grade.Fields("LRN").Value & "' AND subject_code = '" & lbl_subject_code.Caption & "'AND Period='3rd Grading'")
        gradeform.txt_3rd.Text = rs_3rd.Fields("grade").Value
        gradeform.lbl_remark_3rd.Text = rs_3rd.Fields("remark").Value
        Call mysql_select(rs_4th, "SELECT * FROM tbl_grade WHERE ID='" & rs_grade.Fields("LRN").Value & "' AND subject_code = '" & lbl_subject_code.Caption & "'AND Period='4th Grading'")
        gradeform.txt_4th.Text = rs_4th.Fields("grade").Value
        gradeform.lbl_remark_4th.Text = rs_4th.Fields("remark").Value
        Call mysql_select(rs_final, "SELECT * FROM tbl_grade WHERE ID='" & rs_grade.Fields("LRN").Value & "' AND subject_code = '" & lbl_subject_code.Caption & "'AND Period='Final'")
        gradeform.lbl_final.Text = rs_final.Fields("grade").Value
        gradeform.lbl_remark_final.Text = rs_final.Fields("remark").Value
    End If
     Call load_form(gradeform, True)

    
End Sub

Private Sub Form_Load()
     Call set_datagrid(dg_grades, rs_grade, _
                                        "SELECT " _
                                            & "a.student_id as LRN, a.last_name as Last_Name, a.First_Name, a.Middle_Name FROM tbl_student a LEFT JOIN tbl_student_level b ON a.student_id = b.ID WHERE b.section_name = '" & section & "' ORDER BY a.Last_Name ASC")
         
End Sub
