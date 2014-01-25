VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form characterencodeform 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Encode Character Grade"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "characterencodeform.frx":0000
   ScaleHeight     =   6000
   ScaleWidth      =   8970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   -120
      Width           =   8775
      Begin VB.ComboBox cmb_period 
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
         ItemData        =   "characterencodeform.frx":1B3CE
         Left            =   4320
         List            =   "characterencodeform.frx":1B3E1
         TabIndex        =   1
         Top             =   600
         Width           =   2535
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
         TabIndex        =   7
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
         TabIndex        =   6
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
         TabIndex        =   5
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
         TabIndex        =   4
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Choose Period:"
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
         TabIndex        =   3
         Top             =   600
         Width           =   1695
      End
   End
   Begin MSDataGridLib.DataGrid dg_grades 
      Height          =   4095
      Left            =   240
      TabIndex        =   8
      Top             =   1320
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
   Begin VB.Label lbl_view_attendance 
      BackStyle       =   0  'Transparent
      Caption         =   "View student's attendance report."
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
      Left            =   5280
      TabIndex        =   2
      Top             =   5520
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Double-click to encode the character grade."
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
      TabIndex        =   9
      Top             =   1080
      Width           =   4695
   End
End
Attribute VB_Name = "characterencodeform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs_grade As New ADODB.Recordset
Public rs_grade2 As New ADODB.Recordset
Dim sql_string As String

Public Sub cmb_period_Change()
   
End Sub

Public Sub cmb_period_Click()
     If cmb_period.Text = "" Then
       Call set_datagrid(dg_grades, rs_grade, _
                                        "SELECT " _
                                            & "a.student_id as LRN, a.last_name as Last_Name, a.First_Name FROM tbl_student a LEFT JOIN tbl_student_level b ON a.student_id = b.ID WHERE b.section_name = '" & section & "'")
        
    Else
            Call mysql_select(public_rs, "SELECT * FROM tbl_character_grade WHERE ID = '" & rs_grade.Fields("LRN") & "' AND Period='" & cmb_period.Text & "'")
            If public_rs.RecordCount = 0 Then
                Call mysql_select(public_rs, "SELECT a.student_id as LRN, a.last_name as Last_Name, a.First_Name FROM tbl_student a LEFT JOIN tbl_student_level b ON a.student_id = b.ID WHERE b.section_name = '" & section & "'")
                While Not public_rs.EOF
                     sql_string = "INSERT INTO " _
                                & "tbl_character_grade (ID, section_name,Period, Honesty,Courtesy,Helpfulness_and_Cooperation,Resourcefulness_and_Creativity,Consideration_for_Others,Sportsmanship,Obedience,Self_Reliance,Industry,Cleanliness_and_Orderliness,Promptness_and_Punctuality,Sense_of_Responsibility,Love_of_God,Patriotism_and_Love_of_Country)" _
                            & " VALUES ('" _
                                 & public_rs.Fields("LRN") & "','" & lbl_section.Caption & "','" _
                                 & cmb_period.Text & "','','','','','','','','','','','','','','')"
                    Call mysql_select(characterencodeform.rs_grade2, sql_string)
                    public_rs.MoveNext
                Wend
                 Call set_datagrid(dg_grades, rs_grade, _
                                            "SELECT " _
                                                & "a.student_id as LRN, a.last_name as Last_Name, a.First_Name, b.Honesty,b.Courtesy,b.Helpfulness_and_Cooperation,b.Resourcefulness_and_Creativity,b.Consideration_for_Others,b.Sportsmanship,b.Obedience,b.Self_Reliance,b.Industry,b.Cleanliness_and_Orderliness,b.Promptness_and_Punctuality,b.Sense_of_Responsibility,b.Love_of_God,b.Patriotism_and_Love_of_Country  FROM tbl_student a LEFT JOIN tbl_character_grade b ON a.student_id = b.ID WHERE b.section_name = '" & section & "' AND b.period = '" & cmb_period.Text & "'")
             
            Else
                Call set_datagrid(dg_grades, rs_grade, _
                                            "SELECT " _
                                                & "a.student_id as LRN, a.last_name as Last_Name, a.First_Name, b.Honesty,b.Courtesy,b.Helpfulness_and_Cooperation,b.Resourcefulness_and_Creativity,b.Consideration_for_Others,b.Sportsmanship,b.Obedience,b.Self_Reliance,b.Industry,b.Cleanliness_and_Orderliness,b.Promptness_and_Punctuality,b.Sense_of_Responsibility,b.Love_of_God,b.Patriotism_and_Love_of_Country  FROM tbl_student a LEFT JOIN tbl_character_grade b ON a.student_id = b.ID WHERE b.section_name = '" & section & "' AND b.period = '" & cmb_period.Text & "'")
             
            End If
            If cmb_period.Text = "Final" Then
                Call mysql_select(public_rs, "SELECT * FROM tbl_attendance WHERE ID = '" & rs_grade.Fields("LRN") & "' AND section_name='" & section & "'")
                If public_rs.RecordCount = 0 Then
                    lbl_view_attendance.Visible = False
                Else
                    lbl_view_attendance.Visible = True
                End If
            Else
                lbl_view_attendance.Visible = False
            End If
      
            
    End If
End Sub

Private Sub cmb_period_KeyUp(KeyCode As Integer, Shift As Integer)
    MsgBox "Please select period from the list."
    cmb_period.Text = ""
End Sub

Private Sub dg_grades_DblClick()
    If rs_grade.RecordCount = 0 Then
        MsgBox "No student enrolled in this section."
    Else
         If cmb_period.Text = "" Then
        MsgBox "Please select a period first."
    Else
        charactergradeform.lbl_id2.Caption = rs_grade.Fields("LRN")
        charactergradeform.lbl_name2.Caption = rs_grade.Fields("Last_Name") & ", " & rs_grade.Fields("First_Name")
        charactergradeform.lbl_period2.Caption = cmb_period.Text
        charactergradeform.cmb_1.Text = rs_grade.Fields("Honesty")
        charactergradeform.cmb_2.Text = rs_grade.Fields("Courtesy")
        charactergradeform.cmb_3.Text = rs_grade.Fields("Helpfulness_and_Cooperation")
        charactergradeform.cmb_4.Text = rs_grade.Fields("Resourcefulness_and_Creativity")
        charactergradeform.cmb_5.Text = rs_grade.Fields("Consideration_for_Others")
        charactergradeform.cmb_6.Text = rs_grade.Fields("Sportsmanship")
        charactergradeform.cmb_7.Text = rs_grade.Fields("Obedience")
        charactergradeform.cmb_8.Text = rs_grade.Fields("Self_Reliance")
        charactergradeform.cmb_9.Text = rs_grade.Fields("Industry")
        charactergradeform.cmb_10.Text = rs_grade.Fields("Cleanliness_and_Orderliness")
        charactergradeform.cmb_11.Text = rs_grade.Fields("Promptness_and_Punctuality")
        charactergradeform.cmb_12.Text = rs_grade.Fields("Sense_of_Responsibility")
        charactergradeform.cmb_13.Text = rs_grade.Fields("Love_of_God")
        charactergradeform.cmb_14.Text = rs_grade.Fields("Patriotism_and_Love_of_Country")
        Call load_form(charactergradeform, True)
    End If
    End If
   
End Sub

Private Sub Form_Load()
     Call set_datagrid(dg_grades, rs_grade, _
                                        "SELECT " _
                                            & "a.student_id as LRN, a.last_name as Last_Name, a.First_Name FROM tbl_student a LEFT JOIN tbl_student_level b ON a.student_id = b.ID WHERE  b.section_name = '" & section & "'")
        
End Sub

Private Sub lbl_view_attendance_Click()
    attendanceform.lbl_id2.Caption = rs_grade.Fields("LRN")
    attendanceform.lbl_name2.Caption = rs_grade.Fields("Last_Name") & ", " & rs_grade.Fields("First_Name")
     Call mysql_select(public_rs, "SELECT * FROM tbl_attendance WHERE ID = '" & rs_grade.Fields("LRN") & "'")
     If public_rs.RecordCount = 0 Then
        MsgBox "No attendance record for this student. Please complete first the final grade of character building."
        Exit Sub
    Else
     attendanceform.txt_school_days.Text = public_rs.Fields("no_school_days")
     attendanceform.txt_days_absent.Text = public_rs.Fields("no_days_absent")
     attendanceform.txt_days_tardy.Text = public_rs.Fields("no_days_tardiness")
     attendanceform.txt_days_present.Text = public_rs.Fields("no_days_present")
    Call load_form(attendanceform, True)
    End If
End Sub
