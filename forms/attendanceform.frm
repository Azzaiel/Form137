VERSION 5.00
Begin VB.Form attendanceform 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Encode Student's Attendance Report"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "attendanceform.frx":0000
   ScaleHeight     =   5985
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt_days_present 
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
      Height          =   495
      Left            =   4560
      TabIndex        =   30
      Text            =   "0"
      Top             =   5400
      Width           =   735
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00808080&
      Caption         =   "Tardiness"
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
      Height          =   3855
      Left            =   4560
      TabIndex        =   28
      Top             =   1320
      Width           =   4335
      Begin VB.TextBox txt_tardy_3 
         Enabled         =   0   'False
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
         TabIndex        =   14
         Top             =   3360
         Width           =   4095
      End
      Begin VB.CheckBox chk_tardy_3 
         BackColor       =   &H00808080&
         Caption         =   "Cause 3"
         Enabled         =   0   'False
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
         Left            =   120
         TabIndex        =   13
         Top             =   3000
         Width           =   1815
      End
      Begin VB.TextBox txt_tardy_2 
         Enabled         =   0   'False
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
         TabIndex        =   12
         Top             =   2520
         Width           =   4095
      End
      Begin VB.CheckBox chk_tardy_2 
         BackColor       =   &H00808080&
         Caption         =   "Cause 2"
         Enabled         =   0   'False
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
         Left            =   120
         TabIndex        =   11
         Top             =   2160
         Width           =   1815
      End
      Begin VB.TextBox txt_tardy_1 
         Enabled         =   0   'False
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
         TabIndex        =   10
         Top             =   1680
         Width           =   4095
      End
      Begin VB.CheckBox chk_tardy_1 
         BackColor       =   &H00808080&
         Caption         =   "Cause 1"
         Enabled         =   0   'False
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
         Left            =   120
         TabIndex        =   9
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox txt_days_tardy 
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
         Left            =   3480
         MaxLength       =   3
         TabIndex        =   8
         Text            =   "0"
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Note: When tardiness is higher than 3, input atleast 3 general causes."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   480
         TabIndex        =   34
         Top             =   720
         Width           =   3495
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Number of times tardy:"
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
         TabIndex        =   32
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      Caption         =   "Absences"
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
      Height          =   3855
      Left            =   120
      TabIndex        =   27
      Top             =   1320
      Width           =   4335
      Begin VB.TextBox txt_absent_3 
         Enabled         =   0   'False
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
         TabIndex        =   7
         Top             =   3360
         Width           =   4095
      End
      Begin VB.CheckBox chk_absent_3 
         BackColor       =   &H00808080&
         Caption         =   "Cause 3"
         Enabled         =   0   'False
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
         Left            =   120
         TabIndex        =   6
         Top             =   3000
         Width           =   1815
      End
      Begin VB.TextBox txt_absent_2 
         Enabled         =   0   'False
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
         TabIndex        =   5
         Top             =   2520
         Width           =   4095
      End
      Begin VB.CheckBox chk_absent_2 
         BackColor       =   &H00808080&
         Caption         =   "Cause 2"
         Enabled         =   0   'False
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
         Left            =   120
         TabIndex        =   4
         Top             =   2160
         Width           =   1815
      End
      Begin VB.TextBox txt_absent_1 
         Enabled         =   0   'False
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
         TabIndex        =   3
         Top             =   1680
         Width           =   4095
      End
      Begin VB.CheckBox chk_absent_1 
         BackColor       =   &H00808080&
         Caption         =   "Cause 1"
         Enabled         =   0   'False
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
         Left            =   120
         TabIndex        =   2
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox txt_days_absent 
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
         Left            =   3480
         MaxLength       =   3
         TabIndex        =   1
         Text            =   "0"
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Note: When absences is higher than 3, input atleast 3 general causes."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   33
         Top             =   720
         Width           =   3735
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Number of school days absent:"
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
         TabIndex        =   31
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.TextBox txt_school_days 
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
      Left            =   4560
      MaxLength       =   3
      TabIndex        =   0
      Text            =   "0"
      Top             =   600
      Width           =   735
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Height          =   615
      Left            =   0
      TabIndex        =   15
      Top             =   -120
      Width           =   9015
      Begin VB.Label lbl_period 
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
         Left            =   2040
         TabIndex        =   25
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label lbl_sub_name 
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
         Left            =   2040
         TabIndex        =   24
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label lbl_code 
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
         Left            =   2040
         TabIndex        =   23
         Top             =   960
         Width           =   735
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
         Left            =   2040
         TabIndex        =   22
         Top             =   600
         Width           =   735
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
         Left            =   120
         TabIndex        =   21
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
         Left            =   2640
         TabIndex        =   20
         Top             =   240
         Width           =   855
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
         TabIndex        =   19
         Top             =   600
         Width           =   3375
      End
      Begin VB.Label lbl_id2 
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
         Left            =   720
         TabIndex        =   18
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label lbl_name2 
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
         Left            =   3480
         TabIndex        =   17
         Top             =   240
         Width           =   5415
      End
   End
   Begin VB.CommandButton cmd_save 
      Height          =   615
      Left            =   7800
      Picture         =   "attendanceform.frx":1B3CE
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Number of school days present:"
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
      Left            =   1200
      TabIndex        =   29
      Top             =   5520
      Width           =   3375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Number of school days:"
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
      Left            =   2040
      TabIndex        =   26
      Top             =   840
      Width           =   2535
   End
End
Attribute VB_Name = "attendanceform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rs_grade As New ADODB.Recordset
Dim sql_string As String
Dim school_days, absences, tardiness, present As Integer

Private Sub chk_absent_1_Click()
    If chk_absent_1 = vbChecked Then
        txt_absent_1.Enabled = True
    Else
        txt_absent_1.Enabled = False
    End If
End Sub

Private Sub chk_absent_2_Click()
    If chk_absent_2 = vbChecked Then
        txt_absent_2.Enabled = True
    Else
        txt_absent_2.Enabled = False
    End If
End Sub

Private Sub chk_absent_3_Click()
      If chk_absent_3 = vbChecked Then
        txt_absent_3.Enabled = True
    Else
        txt_absent_3.Enabled = False
    End If
End Sub

Private Sub chk_tardy_1_Click()
     If chk_tardy_1 = vbChecked Then
        txt_tardy_1.Enabled = True
    Else
        txt_tardy_1.Enabled = False
    End If
End Sub

Private Sub chk_tardy_2_Click()
     If chk_tardy_2 = vbChecked Then
        txt_tardy_2.Enabled = True
    Else
        txt_tardy_2.Enabled = False
    End If
End Sub

Private Sub chk_tardy_3_Click()
     If chk_tardy_3 = vbChecked Then
        txt_tardy_3.Enabled = True
    Else
        txt_tardy_3.Enabled = False
    End If
End Sub

Private Sub cmd_cancel_Click()

End Sub

Private Sub cmd_save_Click()
    Dim ans As String
    Dim causes_absences, causes_tardiness As String
    Call mysql_select(public_rs, "SELECT * FROM tbl_attendance WHERE ID = '" & lbl_id2.Caption & "'")
    causes_absences = txt_absent_1.Text & " " & txt_absent_2.Text & " " & txt_absent_3.Text
    causes_tardiness = txt_tardy_1.Text & " " & txt_tardy_2.Text & " " & txt_tardy_3.Text
    If public_rs.RecordCount = 0 Then
        ans = MsgBox("Are you sure you want to save student's attendance report?", vbYesNo, "Attendance Report")
                    If ans = vbNo Then
                        Exit Sub
                    Else
         sql_string = "INSERT INTO " _
                                & "tbl_attendance ( SY, ID, section_name, no_school_days,no_days_absent, causes_of_absences,no_days_tardiness,causes_of_tardiness,no_days_present)" _
                            & " VALUES ('" & mainteacherform.cmb_sy.Text & "', '" _
                                & lbl_id2.Caption & "','" _
                                & masterlistadvisoriesform.lbl_section & "','" & txt_school_days.Text & "','" & txt_days_absent.Text & "','" & causes_absences & "','" & txt_days_tardy.Text & "','" & causes_tardiness & "','" & txt_days_present.Text & "')"
        Call mysql_select(attendanceform.rs_grade, sql_string)
        MsgBox "Attendance report has been added."
        Call load_grade
        End If
    Else
         ans = MsgBox("Are you sure you want to update student's attendance report?", vbYesNo, "Attendance Report")
                    If ans = vbNo Then
                        Exit Sub
                    Else
         sql_string = "UPDATE tbl_attendance SET no_school_days='" & txt_school_days.Text & "', no_days_absent='" & txt_days_absent.Text & "',causes_of_absences='" & causes_absences & "',no_days_tardiness='" & txt_days_tardy.Text & "', causes_of_tardiness = '" & causes_tardiness & "', no_days_present = '" & txt_days_present.Text & "' WHERE ID= '" & lbl_id2.Caption & "'"
        Call mysql_select(attendanceform.rs_grade, sql_string)
        MsgBox "Attendance report has been updated."
        Call load_grade
        End If
    End If
   
End Sub

Private Sub load_grade()
      Call set_datagrid(characterencodeform.dg_grades, rs_grade, _
                                            "SELECT " _
                                                & "a.student_id as LRN, a.last_name as Last_Name, a.First_Name, b.Honesty,b.Courtesy,b.Helpfulness_and_Cooperation,b.Resourcefulness_and_Creativity,b.Consideration_for_Others,b.Sportsmanship,b.Obedience,b.Self_Reliance,b.Industry,b.Cleanliness_and_Orderliness,b.Promptness_and_Punctuality,b.Sense_of_Responsibility,b.Love_of_God,b.Patriotism_and_Love_of_Country  FROM tbl_student a LEFT JOIN tbl_character_grade b ON a.student_id = b.ID WHERE  b.section_name = '" & section & "' AND b.period = '" & charactergradeform.lbl_period2.Caption & "'")
       Call characterencodeform.cmb_period_Click
       Unload Me
       Unload charactergradeform
       
End Sub

Private Sub txt_days_absent_KeyUp(KeyCode As Integer, Shift As Integer)
If Not IsNumeric(txt_days_absent.Text) Then
    
     MsgBox "Please enter numbers only."
     txt_days_absent.Text = ""
     Exit Sub
     End If
        absences = val(txt_days_absent.Text)

    school_days = val(txt_school_days)
    
    present = school_days - absences
    txt_days_present.Text = present
    If absences = 0 Then
        chk_absent_1.Enabled = False
        chk_absent_2.Enabled = False
        chk_absent_3.Enabled = False
        txt_absent_1.Enabled = False
        txt_absent_2.Enabled = False
        txt_absent_3.Enabled = False
    ElseIf absences = 1 Then
        chk_absent_1.Enabled = True
        chk_absent_2.Enabled = False
        chk_absent_3.Enabled = False
        txt_absent_1.Enabled = False
        txt_absent_2.Enabled = False
        txt_absent_3.Enabled = False
    ElseIf absences = 2 Then
        chk_absent_1.Enabled = True
        chk_absent_2.Enabled = True
        chk_absent_3.Enabled = False
        txt_absent_1.Enabled = False
        txt_absent_2.Enabled = False
        txt_absent_3.Enabled = False
    ElseIf absences > 2 Then
        chk_absent_1.Enabled = True
        chk_absent_2.Enabled = True
        chk_absent_3.Enabled = True
        txt_absent_1.Enabled = False
        txt_absent_2.Enabled = False
        txt_absent_3.Enabled = False
    End If

End Sub

Private Sub txt_days_tardy_KeyUp(KeyCode As Integer, Shift As Integer)
    If Not IsNumeric(txt_days_tardy.Text) Then
    
     MsgBox "Please enter numbers only."
     txt_days_tardy.Text = ""
     Exit Sub
     End If
    tardiness = val(txt_days_tardy.Text)
    If tardiness = 0 Then
        chk_tardy_1.Enabled = False
        chk_tardy_2.Enabled = False
        chk_tardy_3.Enabled = False
        txt_tardy_1.Enabled = False
        txt_tardy_2.Enabled = False
        txt_tardy_3.Enabled = False
    ElseIf tardiness = 1 Then
        chk_tardy_1.Enabled = True
        chk_tardy_2.Enabled = False
        chk_tardy_3.Enabled = False
        txt_tardy_1.Enabled = False
        txt_tardy_2.Enabled = False
        txt_tardy_3.Enabled = False
    ElseIf tardiness = 2 Then
        chk_tardy_1.Enabled = True
        chk_tardy_2.Enabled = True
        chk_tardy_3.Enabled = False
        txt_tardy_1.Enabled = False
        txt_tardy_2.Enabled = False
        txt_tardy_3.Enabled = False
    ElseIf tardiness > 2 Then
        chk_tardy_1.Enabled = True
        chk_tardy_2.Enabled = True
        chk_tardy_3.Enabled = True
        txt_tardy_1.Enabled = False
        txt_tardy_2.Enabled = False
        txt_tardy_3.Enabled = False
    End If
End Sub

Private Sub txt_school_days_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If Not IsNumeric(txt_school_days.Text) Then
    
     MsgBox "Please enter numbers only."
     txt_school_days.Text = ""
     Exit Sub
     End If
    school_days = val(txt_school_days)
    absences = val(txt_days_absent.Text)
    present = school_days - absences
    txt_days_present.Text = present
End Sub

