VERSION 5.00
Begin VB.Form gradeform 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Grade"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "gradeform.frx":0000
   ScaleHeight     =   6120
   ScaleWidth      =   8865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00808080&
      Caption         =   "Legend"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   960
      TabIndex        =   30
      Top             =   1680
      Width           =   6855
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "74% and below"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   480
         TabIndex        =   40
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "B - Beginning"
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
         Height          =   375
         Left            =   240
         TabIndex        =   39
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "75% - 79%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   2520
         TabIndex        =   38
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "D - Developing"
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
         Height          =   375
         Left            =   2280
         TabIndex        =   37
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "80% - 84%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   4680
         TabIndex        =   36
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "AP - Approaching Proficiency"
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
         Height          =   495
         Left            =   4320
         TabIndex        =   35
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "85% - 89%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   2640
         TabIndex        =   34
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "P - Proficient"
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
         Height          =   375
         Left            =   2280
         TabIndex        =   33
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "90% and above"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   600
         TabIndex        =   32
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "A - Advanced"
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
         Height          =   375
         Left            =   240
         TabIndex        =   31
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.TextBox lbl_remark_final 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   7680
      TabIndex        =   27
      Top             =   4920
      Width           =   855
   End
   Begin VB.TextBox lbl_remark_4th 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   6000
      TabIndex        =   26
      Top             =   4920
      Width           =   855
   End
   Begin VB.TextBox lbl_remark_3rd 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   4440
      TabIndex        =   25
      Top             =   4920
      Width           =   855
   End
   Begin VB.TextBox lbl_remark_2nd 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2880
      TabIndex        =   24
      Top             =   4920
      Width           =   855
   End
   Begin VB.TextBox lbl_remark_1st 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1320
      TabIndex        =   23
      Top             =   4920
      Width           =   855
   End
   Begin VB.TextBox lbl_final 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   7680
      TabIndex        =   21
      Text            =   "0"
      Top             =   4080
      Width           =   855
   End
   Begin VB.TextBox txt_4th 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      MaxLength       =   3
      TabIndex        =   4
      Text            =   "0"
      Top             =   4080
      Width           =   855
   End
   Begin VB.TextBox txt_3rd 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      MaxLength       =   3
      TabIndex        =   3
      Text            =   "0"
      Top             =   4080
      Width           =   855
   End
   Begin VB.TextBox txt_2nd 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      MaxLength       =   3
      TabIndex        =   2
      Text            =   "0"
      Top             =   4080
      Width           =   855
   End
   Begin VB.CommandButton cmd_save 
      Height          =   615
      Left            =   7440
      Picture         =   "gradeform.frx":1B3CE
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5520
      Width           =   1095
   End
   Begin VB.TextBox txt_1st 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      MaxLength       =   3
      TabIndex        =   1
      Text            =   "0"
      Top             =   4080
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   8895
      Begin VB.Label lbl_subject_name2 
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
         Left            =   2880
         TabIndex        =   16
         Top             =   1320
         Width           =   5655
      End
      Begin VB.Label lbl_section2 
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
         Left            =   2880
         TabIndex        =   15
         Top             =   960
         Width           =   3855
      End
      Begin VB.Label lbl_name 
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
         Left            =   2880
         TabIndex        =   14
         Top             =   600
         Width           =   3855
      End
      Begin VB.Label lbl_id 
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
         Left            =   2880
         TabIndex        =   13
         Top             =   240
         Width           =   5775
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Subject Code:"
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
         Left            =   960
         TabIndex        =   11
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Section"
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
         Left            =   960
         TabIndex        =   10
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "LRN:"
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
         Left            =   960
         TabIndex        =   9
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label3 
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
         Left            =   960
         TabIndex        =   8
         Top             =   600
         Width           =   855
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
         TabIndex        =   6
         Top             =   600
         Width           =   3375
      End
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Remark"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   29
      Top             =   9000
      Width           =   975
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Grade"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   28
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Final"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7800
      TabIndex        =   22
      Top             =   3720
      Width           =   735
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "4th Grading"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   20
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "3rd Grading"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   19
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "2nd Grading"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   18
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "1st Grading"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   17
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Note: Input a final grade per period."
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
      Left            =   240
      TabIndex        =   12
      Top             =   3360
      Width           =   5055
   End
End
Attribute VB_Name = "gradeform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rs_grade As New ADODB.Recordset
Dim sql_string As String
Public rs_1st As New ADODB.Recordset
Public rs_2nd As New ADODB.Recordset
Public rs_3rd As New ADODB.Recordset
Public rs_4th As New ADODB.Recordset
Public rs_final As New ADODB.Recordset
Dim first, second, third, fourth, final As Double
Private Sub cmd_save_Click()
    Dim ans As String
   Call mysql_select(public_rs, "SELECT * FROM tbl_grade WHERE ID='" & lbl_id.Caption & "' AND subject_code = '" & lbl_subject_name2.Caption & "'")
    If public_rs.RecordCount = 0 Then
                 ans = MsgBox("Are you sure you want to save the student's grades?", vbYesNo, "Add Grades")
                    If ans = vbNo Then
                        Exit Sub
                    Else
                 sql_string = "INSERT INTO " _
                                & "tbl_grade (ID,section_name,subject_code,period,grade, remark)" _
                            & " VALUES ('" _
                                & lbl_id.Caption & "','" _
                                & lbl_section2.Caption & "','" & lbl_subject_name2.Caption & "','1st Grading','" & txt_1st.Text & "','" & lbl_remark_1st.Text & "')"
              
                Call mysql_select(rs_1st, sql_string)
                 sql_string = "INSERT INTO " _
                                & "tbl_grade (ID,section_name,subject_code,period,grade, remark)" _
                            & " VALUES ('" _
                                & lbl_id.Caption & "','" _
                                & lbl_section2.Caption & "','" & lbl_subject_name2.Caption & "','2nd Grading','" & txt_2nd.Text & "','" & lbl_remark_2nd.Text & "')"
              
                Call mysql_select(rs_2nd, sql_string)
                 sql_string = "INSERT INTO " _
                                & "tbl_grade (ID,section_name,subject_code,period,grade, remark)" _
                            & " VALUES ('" _
                                & lbl_id.Caption & "','" _
                                & lbl_section2.Caption & "','" & lbl_subject_name2.Caption & "','3rd Grading','" & txt_3rd.Text & "','" & lbl_remark_3rd.Text & "')"
                
                Call mysql_select(rs_3rd, sql_string)
                 sql_string = "INSERT INTO " _
                                & "tbl_grade (ID,section_name,subject_code,period,grade, remark)" _
                            & " VALUES ('" _
                                & lbl_id.Caption & "','" _
                                & lbl_section2.Caption & "','" & lbl_subject_name2.Caption & "','4th Grading','" & txt_4th.Text & "','" & lbl_remark_4th.Text & "')"
                
                Call mysql_select(rs_4th, sql_string)
                 sql_string = "INSERT INTO " _
                                & "tbl_grade (ID,section_name,subject_code,period,grade, remark)" _
                            & " VALUES ('" _
                                & lbl_id.Caption & "','" _
                                & lbl_section2.Caption & "','" & lbl_subject_name2.Caption & "','Final','" & lbl_final.Text & "','" & lbl_remark_final.Text & "')"
                
                Call mysql_select(rs_final, sql_string)
                MsgBox "Student's grades saved."
                End If
    Else
         ans = MsgBox("Are you sure you want to update the student's grades?", vbYesNo, "Update Grades")
                    If ans = vbNo Then
                        Exit Sub
                    Else
         sql_string = "UPDATE tbl_grade SET Grade = '" & txt_1st.Text & "', Remark ='" & lbl_remark_1st.Text & "' WHERE ID='" & lbl_id.Caption & "' AND section_name='" & lbl_section2.Caption & "' AND subject_code ='" & lbl_subject_name2.Caption & "' AND Period='1st Grading'"
        Call mysql_select(rs_1st, sql_string)
        sql_string = "UPDATE tbl_grade SET Grade = '" & txt_2nd.Text & "', Remark ='" & lbl_remark_2nd.Text & "' WHERE  ID='" & lbl_id.Caption & "' AND section_name='" & lbl_section2.Caption & "' AND subject_code='" & lbl_subject_name2.Caption & "' AND Period='2nd Grading'"
        Call mysql_select(rs_2nd, sql_string)
        sql_string = "UPDATE tbl_grade SET Grade = '" & txt_3rd.Text & "', Remark ='" & lbl_remark_3rd.Text & "' WHERE ID='" & lbl_id.Caption & "' AND section_name='" & lbl_section2.Caption & "' AND subject_code='" & lbl_subject_name2.Caption & "' AND Period='3rd Grading'"
        Call mysql_select(rs_3rd, sql_string)
        sql_string = "UPDATE tbl_grade SET Grade = '" & txt_4th.Text & "', Remark ='" & lbl_remark_4th.Text & "' WHERE ID='" & lbl_id.Caption & "' AND section_name='" & lbl_section2.Caption & "' AND subject_code='" & lbl_subject_name2.Caption & "' AND Period='4th Grading'"
        Call mysql_select(rs_4th, sql_string)
        sql_string = "UPDATE tbl_grade SET Grade = '" & lbl_final.Text & "', Remark ='" & lbl_remark_final.Text & "' WHERE ID='" & lbl_id.Caption & "' AND section_name='" & lbl_section2.Caption & "' AND subject_code='" & lbl_subject_name2.Caption & "' AND Period='Final'"
        Call mysql_select(rs_final, sql_string)
        MsgBox "Student's grades updated."
        End If
    End If
                
End Sub

Private Sub txt_grade_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim grade, no As Integer
    
        
        
     On Error GoTo message
        grade = Int(txt_grade.Text)
        'no = grade.length
        'If no = 2 Then
            'If grade < 50 Then
                'MsgBox "Please input a grade at least 50."
                'txt_grade.Text = "0"
            'End If
        'ElseIf no > 2 Then
            'If grade > 100 Then
                'MsgBox "Please input a grade at most 50."
                'txt_grade.Text = "0"
            'End If
        'End If
        Exit Sub
message:
        MsgBox "Please input a number."
        txt_grade.Text = "0"
    
End Sub
Private Sub load_grade()
    Call set_datagrid(gradeencodeform.dg_grades, rs_grade, _
                                            "SELECT " _
                                                & "a.student_id as LRN, a.last_name as Last_Name, a.First_Name, b.grade as Grade FROM tbl_student a LEFT JOIN tbl_grade b ON a.student_id = b.ID WHERE b.SY='" & mainteacherform.lbl_sy.Caption & "' AND b.section_name = '" & section & "' AND b.period = '" & lbl_period.Caption & "'")
     Unload Me
End Sub

Private Sub txt_1st_KeyUp(KeyCode As Integer, Shift As Integer)
     If Not IsNumeric(txt_1st.Text) Then
    
        MsgBox "Please enter numbers only."
        txt_1st.Text = ""
        Exit Sub
     End If
     first = val(txt_1st.Text)
     If first > 100 Then
        MsgBox "Invalid grade."
        txt_1st.Text = "0"
        first = 0
     End If
    If first > 100 Then
        lbl_remark_1st.Text = "Invalid"
    ElseIf first > 89 Then
        lbl_remark_1st.Text = "A"
    ElseIf first > 84 Then
        lbl_remark_1st.Text = "P"
    ElseIf first > 79 Then
        lbl_remark_1st.Text = "AP"
    ElseIf first > 74 Then
        lbl_remark_1st.Text = "D"
    ElseIf first > 49 Then
        lbl_remark_1st.Text = "B"
    Else
        lbl_remark_1st.Text = "Invalid"
    End If
    final = (first + second + third + fourth) / 4
    final = Round(final, 2)
    lbl_final.Text = final
    Call remark
End Sub

Private Sub txt_2nd_KeyUp(KeyCode As Integer, Shift As Integer)
     If Not IsNumeric(txt_2nd.Text) Then
    
        MsgBox "Please enter numbers only."
        txt_2nd.Text = ""
        Exit Sub
     End If
      second = val(txt_2nd.Text)
       If second > 100 Then
        MsgBox "Invalid grade."
        txt_2nd.Text = "0"
        second = 0
     End If
    If second > 100 Then
        lbl_remark_2nd.Text = "Invalid"
    ElseIf second > 89 Then
        lbl_remark_2nd.Text = "A"
    ElseIf second > 84 Then
        lbl_remark_2nd.Text = "P"
    ElseIf second > 79 Then
        lbl_remark_2nd.Text = "AP"
    ElseIf second > 74 Then
        lbl_remark_2nd.Text = "D"
    ElseIf second > 49 Then
        lbl_remark_2nd.Text = "B"
    Else
        lbl_remark_2nd.Text = "Invalid"
    End If
    final = (first + second + third + fourth) / 4
    final = Round(final, 2)
    lbl_final.Text = final
    Call remark
End Sub

Private Sub txt_3rd_KeyUp(KeyCode As Integer, Shift As Integer)
     If Not IsNumeric(txt_3rd.Text) Then
    
        MsgBox "Please enter numbers only."
        txt_3rd.Text = ""
        Exit Sub
     End If
       third = val(txt_3rd.Text)
        If third > 100 Then
        MsgBox "Invalid grade."
        txt_3rd.Text = "0"
        third = 0
     End If
    If third > 100 Then
        lbl_remark_3rd.Text = "Invalid"
    ElseIf third > 89 Then
        lbl_remark_3rd.Text = "A"
    ElseIf third > 84 Then
        lbl_remark_3rd.Text = "P"
    ElseIf third > 79 Then
        lbl_remark_3rd.Text = "AP"
    ElseIf third > 74 Then
        lbl_remark_3rd.Text = "D"
    ElseIf third > 49 Then
        lbl_remark_3rd.Text = "B"
    Else
        lbl_remark_3rd.Text = "Invalid"
    End If
    final = (first + second + third + fourth) / 4
    final = Round(final, 2)
    lbl_final.Text = final
    Call remark
End Sub

Public Sub remark()
    If final > 100 Then
        lbl_remark_final.Text = "Invalid"
    ElseIf final > 89 Then
        lbl_remark_final.Text = "A"
    ElseIf final > 84 Then
        lbl_remark_final.Text = "P"
    ElseIf final > 79 Then
        lbl_remark_final.Text = "AP"
    ElseIf final > 74 Then
        lbl_remark_final.Text = "D"
    ElseIf final > 49 Then
        lbl_remark_final.Text = "B"
    Else
        lbl_remark_final.Text = "Incomplete"
    End If
End Sub

Private Sub txt_4th_KeyUp(KeyCode As Integer, Shift As Integer)
     If Not IsNumeric(txt_4th.Text) Then
    
        MsgBox "Please enter numbers only."
        txt_4th.Text = ""
        Exit Sub
     End If
       fourth = val(txt_4th.Text)
        If fourth > 100 Then
        MsgBox "Invalid grade."
        txt_4th.Text = "0"
        fourth = 0
     End If
    If fourth > 100 Then
        lbl_remark_4th.Text = "Invalid"
    ElseIf fourth > 89 Then
        lbl_remark_4th.Text = "A"
    ElseIf fourth > 84 Then
        lbl_remark_4th.Text = "P"
    ElseIf fourth > 79 Then
        lbl_remark_4th.Text = "AP"
    ElseIf fourth > 74 Then
        lbl_remark_4th.Text = "D"
    ElseIf fourth > 49 Then
        lbl_remark_4th.Text = "B"
    Else
        lbl_remark_4th.Text = "Invalid"
    End If
    final = (first + second + third + fourth) / 4
    final = Round(final, 2)
    lbl_final.Text = final
    Call remark
End Sub
