VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form studentinformationform 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Student Information"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "studentinformationform.frx":0000
   ScaleHeight     =   5925
   ScaleWidth      =   8805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt_occupation 
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
      Left            =   1920
      MaxLength       =   12
      TabIndex        =   32
      Top             =   4920
      Width           =   3495
   End
   Begin VB.TextBox txt_place 
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
      Left            =   1920
      MaxLength       =   12
      TabIndex        =   30
      Top             =   3480
      Width           =   3495
   End
   Begin VB.TextBox txt_id2 
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
      Left            =   3720
      MaxLength       =   6
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txt_oldid 
      Height          =   375
      Left            =   1440
      TabIndex        =   27
      Top             =   6000
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox cmb_status 
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
      ItemData        =   "studentinformationform.frx":1B3CE
      Left            =   5640
      List            =   "studentinformationform.frx":1B3D8
      TabIndex        =   10
      Top             =   4440
      Width           =   3015
   End
   Begin VB.TextBox txt_op 
      Height          =   375
      Left            =   120
      TabIndex        =   25
      Top             =   6000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txt_age 
      BackColor       =   &H00E0E0E0&
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
      Left            =   4440
      TabIndex        =   24
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox txt_id 
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
      Left            =   1920
      MaxLength       =   12
      TabIndex        =   28
      Top             =   600
      Width           =   1695
   End
   Begin VB.TextBox txt_lastname 
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
      Left            =   1920
      MaxLength       =   35
      TabIndex        =   1
      Top             =   1080
      Width           =   3495
   End
   Begin VB.TextBox txt_firstname 
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
      Left            =   1920
      MaxLength       =   35
      TabIndex        =   2
      Top             =   1560
      Width           =   3495
   End
   Begin VB.TextBox txt_middlename 
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
      Left            =   1920
      MaxLength       =   35
      TabIndex        =   3
      Top             =   2040
      Width           =   3495
   End
   Begin VB.TextBox txt_no 
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
      Left            =   1920
      TabIndex        =   6
      Top             =   3960
      Width           =   3495
   End
   Begin VB.TextBox txt_address 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5640
      TabIndex        =   9
      Top             =   2520
      Width           =   3015
   End
   Begin VB.TextBox txt_father 
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
      Left            =   1920
      MaxLength       =   50
      TabIndex        =   7
      Top             =   4440
      Width           =   3495
   End
   Begin VB.TextBox txt_father_no 
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
      Left            =   1920
      TabIndex        =   8
      Top             =   5400
      Width           =   3495
   End
   Begin VB.CommandButton cmd_save 
      Height          =   615
      Left            =   6480
      Picture         =   "studentinformationform.frx":1B3F4
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5160
      Width           =   1095
   End
   Begin VB.ComboBox cmb_gender 
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
      ItemData        =   "studentinformationform.frx":1C397
      Left            =   1920
      List            =   "studentinformationform.frx":1C3A1
      TabIndex        =   4
      Top             =   2520
      Width           =   3495
   End
   Begin MSComCtl2.DTPicker dateBday 
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   3000
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   110362625
      CurrentDate     =   41608
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "*Occupation"
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
      Left            =   120
      TabIndex        =   33
      Top             =   5040
      Width           =   1815
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "*Birth Place:"
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
      Left            =   120
      TabIndex        =   31
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Note: Please fill-up important fields."
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
      TabIndex        =   29
      Top             =   120
      Width           =   7335
   End
   Begin VB.Label transferee 
      BackStyle       =   0  'Transparent
      Caption         =   "Click to Input LRN for Transferee"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   6120
      TabIndex        =   12
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label lbl_status 
      BackStyle       =   0  'Transparent
      Caption         =   "*Status:"
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
      Left            =   5640
      TabIndex        =   26
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Age:"
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
      Left            =   3720
      TabIndex        =   23
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "*LRN:"
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
      Left            =   120
      TabIndex        =   22
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "*Last Name:"
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
      Left            =   120
      TabIndex        =   21
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "*First Name:"
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
      Left            =   120
      TabIndex        =   20
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Middle Name:"
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
      Left            =   120
      TabIndex        =   19
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "*Gender:"
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
      Left            =   120
      TabIndex        =   18
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "*Birthday:"
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
      Left            =   120
      TabIndex        =   17
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Number:"
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
      Left            =   120
      TabIndex        =   16
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "*Address:"
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
      Left            =   5640
      TabIndex        =   15
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "*Guardian:"
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
      Left            =   120
      TabIndex        =   14
      Top             =   4560
      Width           =   1815
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "*Number:"
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
      Left            =   120
      TabIndex        =   13
      Top             =   5520
      Width           =   2175
   End
End
Attribute VB_Name = "studentinformationform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rs_student As New ADODB.Recordset
Dim sql_string As String

Private Sub cmb_gender_KeyUp(KeyCode As Integer, Shift As Integer)
    MsgBox "Please select a gender from the list."
    cmb_gender.Text = ""
End Sub

Private Sub cmb_level_Change()
    Call mysql_select(public_rs, "SELECT * FROM tbl_section WHERE SY='" & mainform.lbl_sy.Caption & "' AND lvl_name = '" & cmb_level.Text & "'")
    cmb_section.Clear
    While Not public_rs.EOF
        cmb_section.AddItem (public_rs.Fields("section_name"))
        public_rs.MoveNext
    Wend
End Sub

Private Sub cmb_level_Click()
    Call mysql_select(public_rs, "SELECT * FROM tbl_section WHERE lvl_name = '" & cmb_level.Text & "'")
    cmb_section.Clear
    While Not public_rs.EOF
        cmb_section.AddItem (public_rs.Fields("section_name"))
        public_rs.MoveNext
    Wend
End Sub


Private Sub cmd_cancel_Click()

End Sub

Private Sub cmb_level_KeyUp(KeyCode As Integer, Shift As Integer)
    MsgBox "Please select level from the list."
    cmb_level.Text = ""
End Sub

Private Sub cmb_section_KeyUp(KeyCode As Integer, Shift As Integer)
    MsgBox "Please select section from the list."
    cmb_section.Text = ""
End Sub

Private Sub cmb_status_KeyUp(KeyCode As Integer, Shift As Integer)
    MsgBox "Please select status from the list."
    cmb_status.Text = ""
End Sub

Private Sub cmd_save_Click()
    Dim ans As String
    Dim id As String
    If txt_op.Text = "add" Then
        If txt_id.Text = "" Or txt_lastname.Text = "" Or txt_firstname.Text = "" Or cmb_status.Text = "" Or cmb_gender.ListIndex = -1 Or txt_place = "" Or txt_father = "" Or txt_occupation = "" Or txt_father_no = "" Or txt_address = "" Or cmb_gender.Text = "" Then
          MsgBox "Please complete important fields."
        Exit Sub
    Else
        Dim no As Integer
            no = Len(txt_id2.Text)
            If no < 6 Then
                MsgBox "Please input only 12 characters for LRN."
                txt_id2.SetFocus
                Exit Sub
    End If
        id = txt_id.Text & txt_id2.Text
        If is_duplicate("tbl_student", "student_id", id) Then
            MsgBox "Student ID already exists."
            Exit Sub
        End If
        
         ans = MsgBox("Are you sure you want to add student?", vbYesNo, "Add Student's Information")
                    If ans = vbNo Then
                        Exit Sub
                    Else
        sql_string = "INSERT INTO " _
                        & "tbl_student (student_id,last_name,first_name,middle_name," _
                        & "gender,bday,birthplace,contact_no,address," _
                        & "guardian,guardian_no, occupation)" _
                    & " VALUES (" _
                        & "'" & id & "','" & txt_lastname.Text & "','" _
                        & txt_firstname.Text & "','" & txt_middlename.Text & "','" _
                        & cmb_gender.Text & "','" & Format(dateBday.value, "yyyy-mm-dd") & "','" & txt_place.Text & "','" _
                        & txt_no.Text & "','" _
                        & txt_address.Text & "', '" & txt_father.Text & "','" & txt_father_no.Text & "','" & txt_occupation.Text & "')"
        Call mysql_select(rs_student, sql_string)
        MsgBox "Student's information successfully added."
        Call studentform.Form_Load
        End If
        End If
    Else
         If txt_id.Text = txt_oldid.Text Then
             If txt_id.Text = "" Or txt_lastname.Text = "" Or txt_firstname.Text = "" Or cmb_status.Text = "" Then
                MsgBox "Please complete all fields."
                Exit Sub
            Else
                 ans = MsgBox("Are you sure you want to update student's information?", vbYesNo, "Update Student's Information")
                    If ans = vbNo Then
                        Exit Sub
                    Else
                sql_string = "UPDATE " _
                                & "tbl_student " _
                            & "SET " _
                                & "last_name = '" & txt_lastname.Text & "'," _
                                & "first_name = '" & txt_firstname.Text & "',middle_name = '" _
                                & txt_middlename.Text & "',gender = '" & cmb_gender.Text & "',bday" _
                                & " = '" & Format(dateBday.value, "yyyy-mm-dd") & "'" _
                                & ",birthplace='" & txt_place.Text & "', contact_no = '" & txt_no.Text _
                                & "',address = '" & txt_address.Text & "',guardian ='" & txt_father.Text & "', guardian_no ='" & txt_father_no.Text & "', occupation='" & txt_occupation.Text & "'" _
                            & "WHERE " _
                                & " student_id = '" & txt_oldid.Text & "'"
                Call mysql_select(rs_student, sql_string)
                Call mysql_select(public_rs, "SELECT * FROM tbl_student_level WHERE ID='" & txt_oldid.Text & "'")
                If public_rs.RecordCount <> 0 Then
                'sql_string = "UPDATE "
                '                & "tbl_student_level " _
                '            & "SET " _
                '                & " lvl_name = '" & cmb_level.Text & "', section_name = '" & cmb_section.Text & "', Status= '" & cmb_status.Text & "'" _
                '            & "WHERE " _
                '                & " ID = '" & txt_oldid.Text & "'"
                Call mysql_select(rs_student, sql_string)
                Else
                    Call mysql_select(public_rs, "SELECT * FROM tbl_student_level WHERE ID='" & txt_oldid.Text & "'AND lvl_name = '" & cmb_level.Text & "'")
                    If public_rs.RecordCount = 0 Then
                    sql_string = "INSERT INTO " _
                        & "tbl_student_level (ID,SY,lvl_name,section_name,Status)" _
                    & " VALUES (" _
                        & "'" _
                        & cmb_level.Text & "','" & cmb_section.Text & "','ENROLLED')"
                        Call mysql_select(rs_student, sql_string)
                    Else
                        MsgBox "Unable to enrol in the same level."
                        Exit Sub
                    End If
                End If
                MsgBox "Student's information updated."
                Call studentform.Form_Load
                End If
            End If
        Else
            If txt_id.Text = "" Or txt_lastname.Text = "" Or txt_firstname.Text = "" Or cmb_level.Text = "" Or cmb_section.Text = "" Or cmb_status.Text = "" Then
                MsgBox "Please complete all fields."
                Exit Sub
            Else
                If is_duplicate("tbl_student", "student_id", txt_id.Text) Then
                    MsgBox "Student ID already exists."
                    Exit Sub
                Else
                 ans = MsgBox("Are you sure you want to update student's information?", vbYesNo, "Update Student's Information")
                    If ans = vbNo Then
                        Exit Sub
                    Else
                 sql_string = "UPDATE " _
                                & "tbl_student " _
                            & "SET " _
                                & "student_id='" & txt_id.Text & "', last_name = '" & txt_lastname.Text & "'," _
                                & "first_name = '" & txt_firstname.Text & "',middle_name = '" _
                                & txt_middlename.Text & "',gender = '" & cmb_gender.Text & "',bday" _
                                & " = '" & Format(dateBday.value, "yyyy-mm-dd") & "'" _
                                & ",birthplace='" & txt_place.Text & "',contact_no = '" & txt_no.Text _
                                & "',address = '" & txt_address.Text & "',guardian ='" & txt_father.Text & "', guardian_no ='" & txt_father_no.Text & "', occupation='" & txt_occupation.Text & "'" _
                            & "WHERE " _
                                & " student_id = '" & txt_oldid.Text & "'"
                Call mysql_select(rs_student, sql_string)
                Call mysql_select(public_rs, "SELECT * FROM tbl_student_level WHERE ID='" & txt_oldid.Text & "' ")
                If public_rs.RecordCount <> 0 Then
                sql_string = "UPDATE " _
                                & "tbl_student_level " _
                            & "SET " _
                                & " ID='" & txt_id.Text & "', lvl_name = '" & cmb_level.Text & "', section_name = '" & cmb_section.Text & "', Status= '" & cmb_status.Text & "'" _
                            & "WHERE " _
                                & " ID = '" & txt_oldid.Text & "'"
                Call mysql_select(rs_student, sql_string)
                Else
                    Call mysql_select(public_rs, "SELECT * FROM tbl_student_level WHERE ID='" & txt_oldid.Text & "'AND lvl_name = '" & cmb_level.Text & "'")
                    If public_rs.RecordCount = 0 Then
                    sql_string = "INSERT INTO " _
                        & "tbl_student_level (ID,lvl_name,section_name,Status)" _
                    & " VALUES (" _
                        & "'" & txt_id.Text & "','" _
                        & cmb_level.Text & "','" & cmb_section.Text & "','ENROLLED')"
                        Call mysql_select(rs_student, sql_string)
                    Else
                        MsgBox "Unable to enrol in the same level."
                        Exit Sub
                    End If
                End If
                 MsgBox "Student's information updated."
               Call studentform.Form_Load
               End If
        End If
    End If
   End If
   End If
    Unload Me
End Sub

Private Sub dateBday_Change()
    age = DateDiff("d", dateBday.value, Date) / 365.25
    age = Round(age * 4, 0) / 4
    txt_age.Text = Str(age)
End Sub


Private Sub Form_Load()
        Call mysql_select(public_rs, "SELECT * FROM tbl_student_level WHERE ID='" & txt_id.Text & "'")
        
        If public_rs.RecordCount <> 0 Then
            section = public_rs.Fields("section_name")
            
            Call mysql_select(public_rs, "SELECT * FROM tbl_section WHERE section_name='" & section & "'")
                level = public_rs.Fields("lvl_name")
            Call mysql_select(public_rs, "SELECT * FROM tbl_level")
            cmb_level.Clear
            While Not public_rs.EOF
                cmb_level.AddItem (public_rs.Fields("lvl_name"))
                public_rs.MoveNext
            Wend
            If Not level = "" Then
                cmb_level.Text = level
                Call cmb_level_Change
            End If
            Call mysql_select(public_rs, "SELECT * FROM tbl_section WHERE lvl_name = '" & cmb_level.Text & "'")
            cmb_section.Clear
            While Not public_rs.EOF
                cmb_section.AddItem (public_rs.Fields("section_name"))
                public_rs.MoveNext
            Wend
            If Not section = "" Then
                cmb_section.Text = section
            End If
        Else
            
            'Call mysql_select(public_rs, "SELECT * FROM tbl_level ")
            'cmb_level.Clear
            'While Not public_rs.EOF
            '    cmb_level.AddItem (public_rs.Fields("lvl_name"))
            '    public_rs.MoveNext
            'Wend
            'If Not level = "" Then
            '    cmb_level.Text = level
            '
             '   Call cmb_level_Change
            'End If
            'Call mysql_select(public_rs, "SELECT * FROM tbl_section WHERE lvl_name = '" & cmb_level.Text & "'")
            'cmb_section.Clear
            'While Not public_rs.EOF
            '    cmb_section.AddItem (public_rs.Fields("section_name"))
            '    public_rs.MoveNext
            'Wend
            'If Not section = "" Then
            '    cmb_section.Text = section
                
           ' End If
        End If
  
    
       

    
End Sub
Private Sub load_student()
     Call set_datagrid(studentform.dg_students, rs_student, _
                                        "SELECT " _
                                            & "student_id as LRN, last_name as Last_Name, first_name as First_Name, middle_name as Middle_Name, Gender as Gender, Bday as Date_Of_Birth, birthplace as Birth_Place, contact_no as Contact_Number,address as Address, Guardian as Guardian, guardian_no as Guardian_Contact, Occupation FROM tbl_student")
                                        
                    
       Unload Me
                                       
End Sub
Function get_File_Ext(file_name As String) As String
    file = Split(file_name, ".")
    get_File_Ext = file(UBound(file))
End Function

Private Sub Text1_Change()

End Sub

Private Sub transferee_Click()
      Call load_form(transfereeform, True)
End Sub

Private Sub txt_father_KeyUp(KeyCode As Integer, Shift As Integer)
     If IsNumeric(Right(txt_father.Text, 1)) = True Then
        txt_father.Text = ""
        MsgBox "Number is not allowed."
    End If
End Sub

Private Sub txt_father_no_KeyPress(KeyAscii As Integer)
   If (Not CommonHelper.isFunctionAscii(KeyAscii) And (Not CommonHelper.isNumberAscii(KeyAscii) And KeyAscii <> 47)) Then
      KeyAscii = 0
      Beep
    End If
End Sub

Private Sub txt_firstname_KeyUp(KeyCode As Integer, Shift As Integer)
     If IsNumeric(Right(txt_firstname.Text, 1)) = True Then
        txt_firstname.Text = ""
        MsgBox "Number is not allowed."
    End If
End Sub

Private Sub txt_id2_KeyUp(KeyCode As Integer, Shift As Integer)
     If Not IsNumeric(txt_id2.Text) Then
    
     MsgBox "Please enter numbers only."
     txt_id2.Text = ""
     Exit Sub
     End If
  
End Sub

Private Sub txt_lastname_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim no As Integer
    
    If txt_op.Text = "add" Then
        no = Len(txt_id2.Text)
        If no < 6 Then
            MsgBox "Please input only 12 characters for LRN."
            
            Exit Sub
        End If
        
    End If
    
     If IsNumeric(Right(txt_lastname.Text, 1)) = True Then
        txt_lastname.Text = ""
        MsgBox "Please input numbers only."
    End If
End Sub

Private Sub txt_middlename_KeyUp(KeyCode As Integer, Shift As Integer)
     If IsNumeric(Right(txt_middlename.Text, 1)) = True Then
        txt_middlename.Text = ""
        MsgBox "Number is not allowed."
    End If
End Sub

Private Sub txt_mother_KeyUp(KeyCode As Integer, Shift As Integer)
     If IsNumeric(Right(txt_mother.Text, 1)) = True Then
        txt_mother.Text = ""
        MsgBox "Number is not allowed."
    End If
End Sub

Private Sub txt_mother_no_KeyUp(KeyCode As Integer, Shift As Integer)
     If Not IsNumeric(txt_mother_no.Text) = True Then
        txt_mother_no.Text = ""
        MsgBox "Please input numbers only."
    End If
End Sub

Private Sub txt_no_KeyPress(KeyAscii As Integer)
    If (Not CommonHelper.isFunctionAscii(KeyAscii) And (Not CommonHelper.isNumberAscii(KeyAscii) And KeyAscii <> 47)) Then
      KeyAscii = 0
      Beep
    End If
End Sub

