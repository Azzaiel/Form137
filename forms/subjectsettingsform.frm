VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form subjectsettingsform 
   BorderStyle     =   0  'None
   Caption         =   "Subject Settings"
   ClientHeight    =   6090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9015
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "subjectsettingsform.frx":0000
   ScaleHeight     =   6090
   ScaleWidth      =   9015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt_code 
      Height          =   375
      Left            =   7680
      TabIndex        =   15
      Top             =   1680
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox cmb_teacher 
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
      Left            =   3840
      TabIndex        =   3
      Top             =   1320
      Width           =   3375
   End
   Begin VB.TextBox txt_subject_name 
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
      Left            =   3840
      MaxLength       =   35
      TabIndex        =   2
      Top             =   840
      Width           =   3375
   End
   Begin VB.CommandButton cmd_save 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      Picture         =   "subjectsettingsform.frx":1B3CE
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton cmb_clear 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4560
      Picture         =   "subjectsettingsform.frx":1C371
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   360
      Width           =   8535
      Begin VB.ComboBox cmb_section 
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
         Left            =   5520
         TabIndex        =   1
         Top             =   0
         Width           =   2895
      End
      Begin VB.ComboBox cmb_level 
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
         Left            =   1560
         TabIndex        =   0
         Top             =   0
         Width           =   2655
      End
      Begin VB.Label Label5 
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
         Left            =   4440
         TabIndex        =   12
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Level Name:"
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
         TabIndex        =   8
         Top             =   120
         Width           =   1455
      End
   End
   Begin MSDataGridLib.DataGrid dg_subjects 
      Height          =   3135
      Left            =   120
      TabIndex        =   6
      Top             =   2880
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   5530
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
   Begin MSDataGridLib.DataGrid dg_subjects2 
      Height          =   3135
      Left            =   4560
      TabIndex        =   13
      Top             =   2880
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   5530
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
   Begin VB.Label lbl_close 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   8640
      TabIndex        =   16
      ToolTipText     =   "Close"
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "List of updated subjects."
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
      Left            =   4560
      TabIndex        =   14
      Top             =   2640
      Width           =   4455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "*Subject Code:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1800
      TabIndex        =   11
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Teacher:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1800
      TabIndex        =   10
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Double click to set teacher per subject."
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
      Left            =   120
      TabIndex        =   9
      Top             =   2640
      Width           =   4455
   End
End
Attribute VB_Name = "subjectsettingsform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs_subject As New ADODB.Recordset
Dim rs_subject2 As New ADODB.Recordset

Private Sub cmb_clear_Click()
    txt_subject_name.Text = ""
    Call mysql_select(public_rs, "SELECT CONCAT(CONCAT(first_name,' '),last_name) as Name FROM tbl_teacher WHERE status = 'Active'")
    cmb_teacher.Clear
    While Not public_rs.EOF
        cmb_teacher.AddItem (public_rs.Fields("Name").Value)
        public_rs.MoveNext
    Wend
End Sub

Private Sub cmb_level_Change()
    txt_subject_name.Text = ""
    cmb_teacher.Clear
    Call set_datagrid(dg_subjects, rs_subject, _
                                        "SELECT " _
                                            & "subject_code as Subject_Code,subject_name as Subject_Name " _
                                        & "FROM " _
                                            & "tbl_subject  " _
                                        & "WHERE " _
                                            & "lvl_name = '" & cmb_level.Text & "' ")
    Call mysql_select(public_rs, "SELECT * FROM tbl_section WHERE lvl_name = '" & cmb_level.Text & "'")
    cmb_section.Clear
    While Not public_rs.EOF
        cmb_section.AddItem (public_rs.Fields("section_name"))
        public_rs.MoveNext
    Wend
End Sub

Private Sub cmb_level_Click()
     txt_subject_name.Text = ""
    cmb_teacher.Clear
    Call set_datagrid(dg_subjects, rs_subject, _
                                        "SELECT " _
                                            & "subject_code as Subject_Code,subject_name as Subject_Name " _
                                        & "FROM " _
                                            & "tbl_subject  ")
     Call mysql_select(public_rs, "SELECT * FROM tbl_section WHERE  lvl_name = '" & cmb_level.Text & "'")
    cmb_section.Clear
    While Not public_rs.EOF
        cmb_section.AddItem (public_rs.Fields("section_name"))
        public_rs.MoveNext
    Wend
End Sub

Private Sub cmb_section_Change()
    txt_subject_name.Text = ""
    cmb_teacher.Clear
    Call set_datagrid(dg_subjects2, rs_subject2, _
                                        "SELECT " _
                                            & "a.subject_code As Subject_Code, a.subject_name As Subject_Name, " _
                                            & "a.teacher_id As Teacher_ID, CONCAT(b.first_name,' ', b.last_name) as Teacher_Name " _
                                        & "FROM " _
                                            & "tbl_subjectset a " _
                                        & "LEFT JOIN " _
                                            & "tbl_teacher b " _
                                        & "ON " _
                                            & "a.teacher_id = b.teacher_id " _
                                        & "WHERE " _
                                            & "a.lvl_name = '" & cmb_level.Text & "' AND a.section_name='" & cmb_section.Text & "'")
     Call mysql_select(public_rs, "SELECT CONCAT(CONCAT(first_name,' '),last_name) as Name FROM tbl_teacher WHERE status = 'Active'")
    cmb_teacher.Clear
    While Not public_rs.EOF
        cmb_teacher.AddItem (public_rs.Fields("Name").Value)
        public_rs.MoveNext
    Wend
End Sub

Private Sub cmb_section_Click()
    txt_subject_name.Text = ""
    cmb_teacher.Clear
    Call set_datagrid(dg_subjects2, rs_subject2, _
                                        "SELECT " _
                                            & "a.subject_code As Subject_Code, a.subject_name As Subject_Name, " _
                                            & "a.teacher_id As Teacher_ID, CONCAT(b.first_name,' ', b.last_name) as Teacher_Name " _
                                        & "FROM " _
                                            & "tbl_subjectset a " _
                                        & "LEFT JOIN " _
                                            & "tbl_teacher b " _
                                        & "ON " _
                                            & "a.teacher_id = b.teacher_id " _
                                        & "WHERE " _
                                            & "a.lvl_name = '" & cmb_level.Text & "' AND a.section_name='" & cmb_section.Text & "'")
     Call mysql_select(public_rs, "SELECT CONCAT(CONCAT(first_name,' '),last_name) as Name FROM tbl_teacher WHERE status = 'Active'")
    cmb_teacher.Clear
    While Not public_rs.EOF
        cmb_teacher.AddItem (public_rs.Fields("Name").Value)
        public_rs.MoveNext
    Wend
End Sub

Private Sub cmb_teacher_KeyUp(KeyCode As Integer, Shift As Integer)
    MsgBox "Please select a teacher from the list."
    cmb_teacher.Text = ""
End Sub

Private Sub cmd_save_Click()
    Dim ans As String
     Dim sql_string As String
    If Not txt_subject_name.Text = "" Then
        If cmb_teacher.Text = "" Then
            MsgBox "Please select a teachet first."
            Exit Sub
        Else
            Call mysql_select(public_rs, "SELECT * FROM tbl_subjectset WHERE lvl_name = '" & cmb_level.Text & "'AND section_name = '" & cmb_section.Text & "'AND subject_code='" & txt_code.Text & "'")
            If public_rs.RecordCount = 0 Then
                Call mysql_select(public_rs, "SELECT teacher_id FROM tbl_teacher WHERE CONCAT(CONCAT(first_name,' '),last_name) = '" & cmb_teacher.Text & "'")
                Dim tch_id As String
                If public_rs.RecordCount = 0 Then
                    tch_id = "None"
                Else
                    tch_id = public_rs.Fields("teacher_id").Value
                End If
                 ans = MsgBox("Are you sure you want to set the teacher for this subject?", vbYesNo, "Set Teacher")
                    If ans = vbNo Then
                        Exit Sub
                    Else
                Call mysql_select(rs_subject2, "INSERT INTO tbl_subjectset(lvl_name,section_name,subject_code,subject_name,teacher_id) VALUES ('" & cmb_level.Text & "','" & cmb_section.Text & "','" & txt_code.Text & "','" & txt_subject_name.Text & "','" & tch_id & "')")
                MsgBox "Teacher for this subject has been set!"
                txt_subject_name.Text = ""
                txt_code.Text = ""
                cmb_teacher.Clear
                level = cmb_level.Text
                section = cmb_section.Text
                Call cmb_section_Click
                End If
            Else
               Call mysql_select(public_rs, "SELECT teacher_id FROM tbl_teacher WHERE CONCAT(CONCAT(first_name,' '),last_name) = '" & cmb_teacher.Text & "'")
                Dim tch_id2 As String
                If public_rs.RecordCount <> 0 Then
                     tch_id2 = public_rs.Fields("teacher_id").Value
                Else
                    tch_id2 = "None"
                End If
                     ans = MsgBox("Are you sure you want to set the teacher for this subject?", vbYesNo, "Set Teacher")
                    If ans = vbNo Then
                        Exit Sub
                    Else
                     sql_string = "UPDATE " _
                            & "tbl_subjectset " _
                        & "SET " _
                            & "teacher_id = '" & tch_id2 & "' " _
                        & "WHERE " _
                            & " lvl_name = '" & cmb_level.Text & "'AND section_name = '" & cmb_section.Text & "'AND subject_code='" & txt_code.Text & "'"
                Call mysql_select(public_rs, sql_string)
                MsgBox "Teacher for this subject has been set!"
                txt_subject_name.Text = ""
                txt_code.Text = ""
                cmb_teacher.Clear
                level = cmb_level.Text
                section = cmb_section.Text
                Call cmb_section_Click
                End If
            End If
         End If
        Else
            MsgBox "Please select a subject."
        End If
   
End Sub

Private Sub dg_subjects_DblClick()
    txt_code.Text = rs_subject.Fields("Subject_Code")
    txt_subject_name.Text = rs_subject.Fields("Subject_Name")
    Call mysql_select(public_rs, "SELECT CONCAT(CONCAT(first_name,' '),last_name) as Name FROM tbl_teacher WHERE status = 'On-Duty'")
    cmb_teacher.Clear
    While Not public_rs.EOF
        cmb_teacher.AddItem (public_rs.Fields("Name").Value)
        public_rs.MoveNext
    Wend
End Sub

Private Sub Form_Load()
    Call mysql_select(public_rs, "SELECT * FROM tbl_level ")
    cmb_level.Clear
    While Not public_rs.EOF
        cmb_level.AddItem (public_rs.Fields("lvl_name"))
        public_rs.MoveNext
    Wend
    If Not level = "" Then
        cmb_level.Text = level
        level = ""
        Call cmb_level_Change
    End If
    Call mysql_select(public_rs, "SELECT * FROM tbl_section WHERE  lvl_name = '" & cmb_level.Text & "'")
    cmb_section.Clear
    While Not public_rs.EOF
        cmb_section.AddItem (public_rs.Fields("section_name"))
        public_rs.MoveNext
    Wend
    If Not section = "" Then
        cmb_section.Text = section
        section = ""
        Call cmb_section_Change
    End If
    
     Call set_datagrid(dg_subjects, rs_subject, _
                                        "SELECT " _
                                            & "subject_code as Subject_Code,subject_name as Subject_Name " _
                                        & "FROM " _
                                            & "tbl_subject  " _
                                        & "WHERE " _
                                            & "lvl_name = '" & cmb_level.Text & "' ")
                                            
    Call set_datagrid(dg_subjects2, rs_subject2, _
                                        "SELECT " _
                                            & "a.subject_code As Subject_Code, a.subject_name As Subject_Name, " _
                                            & "a.teacher_id As Teacher_ID, CONCAT(b.first_name,' ', b.last_name) as Teacher_Name " _
                                        & "FROM " _
                                            & "tbl_subjectset a " _
                                        & "LEFT JOIN " _
                                            & "tbl_teacher b " _
                                        & "ON " _
                                            & "a.teacher_id = b.teacher_id " _
                                        & "WHERE " _
                                            & "a.lvl_name = '" & cmb_level.Text & "' AND a.section_name='" & cmb_section.Text & "'")
End Sub

Private Sub lbl_close_Click()
    Unload Me
    Unload sectionform
    Unload subjectform
    Unload levelform
End Sub
