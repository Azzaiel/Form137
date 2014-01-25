VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form teacherinformationform 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Teacher Information"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "teacherinformationform.frx":0000
   ScaleHeight     =   5985
   ScaleWidth      =   8985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt_oldid 
      Height          =   375
      Left            =   720
      TabIndex        =   29
      Top             =   5520
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txt_op 
      Height          =   375
      Left            =   2520
      TabIndex        =   28
      Top             =   5520
      Visible         =   0   'False
      Width           =   1335
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
      ItemData        =   "teacherinformationform.frx":1B3CE
      Left            =   1920
      List            =   "teacherinformationform.frx":1B3D8
      TabIndex        =   11
      Top             =   5160
      Width           =   2175
   End
   Begin MSComCtl2.DTPicker dateBday 
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   2760
      Width           =   3495
      _ExtentX        =   6165
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
      Format          =   89391105
      CurrentDate     =   41540
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
      ItemData        =   "teacherinformationform.frx":1B3EF
      Left            =   1920
      List            =   "teacherinformationform.frx":1B3F9
      TabIndex        =   4
      Top             =   2280
      Width           =   3495
   End
   Begin VB.CommandButton cmd_save 
      Height          =   615
      Left            =   6000
      Picture         =   "teacherinformationform.frx":1B40B
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton cmd_cancel 
      Height          =   615
      Left            =   7320
      Picture         =   "teacherinformationform.frx":1C3AE
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5280
      Width           =   1095
   End
   Begin VB.TextBox txt_to 
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
      Left            =   6600
      MaxLength       =   4
      TabIndex        =   10
      Top             =   4680
      Width           =   2175
   End
   Begin VB.TextBox txt_from 
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
      MaxLength       =   4
      TabIndex        =   9
      Top             =   4680
      Width           =   2175
   End
   Begin VB.TextBox txt_school 
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
      MaxLength       =   100
      TabIndex        =   8
      Top             =   4200
      Width           =   6855
   End
   Begin VB.TextBox txt_course 
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
      MaxLength       =   100
      TabIndex        =   7
      Top             =   3720
      Width           =   6855
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
      Height          =   1335
      Left            =   5640
      TabIndex        =   12
      Top             =   1320
      Width           =   3015
   End
   Begin VB.TextBox txt_contact 
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
      TabIndex        =   6
      Top             =   3240
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
      Top             =   1800
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
      Top             =   1320
      Width           =   3495
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
      Top             =   840
      Width           =   3495
   End
   Begin VB.TextBox txt_id 
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
      Left            =   1920
      TabIndex        =   0
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label Label14 
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
      Left            =   120
      TabIndex        =   30
      Top             =   0
      Width           =   7335
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Status:"
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
      TabIndex        =   27
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "To:"
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
      Left            =   5040
      TabIndex        =   26
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "From:"
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
      TabIndex        =   25
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "School:"
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
      TabIndex        =   24
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Course:"
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
      TabIndex        =   23
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
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
      TabIndex        =   22
      Top             =   960
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
      TabIndex        =   21
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Birthday:"
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
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Gender:"
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
      Top             =   2400
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
      TabIndex        =   18
      Top             =   1920
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
      TabIndex        =   17
      Top             =   1440
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
      TabIndex        =   16
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Teacher ID:"
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
      TabIndex        =   15
      Top             =   480
      Width           =   1455
   End
End
Attribute VB_Name = "teacherinformationform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rs_teacher As New ADODB.Recordset
Dim sql_string As String
Private Sub cmd_browse_Click()
    On Error GoTo message
        cdPhoto.ShowOpen
        photo.Picture = LoadPicture(cdPhoto.FileName)
      
        Exit Sub
message:
    MsgBox "Problem in loading pictures"
End Sub

Function get_File_Ext(file_name As String) As String
    file = Split(file_name, ".")
    get_File_Ext = file(UBound(file))
End Function

Private Sub cmb_gender_KeyUp(KeyCode As Integer, Shift As Integer)
    MsgBox "Please select gender from the list."
    cmb_gender.Text = ""
End Sub

Private Sub cmb_status_KeyUp(KeyCode As Integer, Shift As Integer)
    MsgBox "Please select status from the list."
    cmb_status.Text = ""
End Sub

Private Sub cmd_cancel_Click()
    txt_id.Text = ""
    txt_firstname.Text = ""
    txt_lastname.Text = ""
    txt_middlename.Text = ""
    cmb_gender.Text = ""
    dateBday.Value = Now
    txt_contact.Text = ""
    txt_address.Text = ""
    txt_course.Text = ""
    txt_school.Text = ""
    txt_from.Text = ""
    txt_to.Text = ""
   cmb_status.Text = "On-Duty"
    txt_op.Text = "add"
    photo.Picture = LoadPicture(App.Path & "\images\photo_teacher\noimage.jpg")
End Sub

Private Sub cmd_save_Click()
    Dim ans As String
    If txt_op.Text = "add" Then
        If txt_id.Text = "" Or txt_lastname.Text = "" Or txt_firstname.Text = "" Then
        MsgBox "Please input required fields."
        Else
        If is_duplicate("tbl_teacher", "teacher_id", txt_id.Text) Then
            MsgBox "Teacher ID already exists."
            Exit Sub
        End If
         ans = MsgBox("Are you sure you want to add teacher's information?", vbYesNo, "Add Teacher")
                    If ans = vbNo Then
                        Exit Sub
                    Else
        sql_string = "INSERT INTO " _
                        & "tbl_teacher (teacher_id,last_name,first_name,middle_name," _
                        & "gender,bday,contact_no,address," _
                        & "course,school,a_from,a_to, status)" _
                    & " VALUES (" _
                        & "'" & txt_id.Text & "','" & txt_lastname.Text & "','" _
                        & txt_firstname.Text & "','" & txt_middlename.Text & "','" _
                        & cmb_gender.Text & "','" & Format(dateBday.Value, "yyyy-mm-dd") & "','" _
                        & txt_contact.Text & "','" _
                        & txt_address.Text & "', '" & txt_course.Text & "','" & txt_school.Text & "','" & txt_from.Text & "','" & txt_to.Text & "', 'On-Duty')"
        Call mysql_select(teacherinformationform.rs_teacher, sql_string)
        Dim USERNAME, lastname As String
        lastname = Replace(txt_lastname.Text, " ", "")
        USERNAME = txt_id.Text & lastname
         sql_string = "INSERT INTO " _
                        & "tbl_user (ID, Usertype, Username, Password)" _
                    & " VALUES (" _
                        & "'" & txt_id.Text & "','Teacher','" _
                        & USERNAME & "','" & lastname & "')"
                        
        Call mysql_select(rs_teacher, sql_string)
        
        Call teacherform.Form_Load
        MsgBox "Teacher's information successfully added."
         Call set_datagrid(teacherform.dg_teachers, rs_teacher, _
                                        "SELECT " _
                                            & "teacher_id as Teacher_ID, last_name as Last_Name, first_name as First_Name, middle_name as Middle_Name, Gender as Gender, Bday as Date_Of_Birth, contact_no as Contact_Number,address as Address, course as Course, school as School_Attended, a_from as From_Year, a_to as To_Year, status as Status FROM tbl_teacher")
                                        
                    
                         
        End If
        End If
        'Call load_teacher
    Else
        If txt_id.Text = txt_oldid.Text Then
            If txt_id.Text = "" Or txt_lastname.Text = "" Or txt_firstname.Text = "" Then
                MsgBox "Please input required fields."
            Else
             ans = MsgBox("Are you sure you want to update teacher's information?", vbYesNo, "Update Teacher")
                    If ans = vbNo Then
                        Exit Sub
                    Else
                sql_string = "UPDATE " _
                                & "tbl_teacher " _
                            & "SET " _
                                & "last_name = '" & txt_lastname.Text & "'," _
                                & "first_name = '" & txt_firstname.Text & "',middle_name = '" _
                                & txt_middlename.Text & "',gender = '" & cmb_gender.Text & "',bday" _
                                & " = '" & Format(dateBday.Value, "yyyy-mm-dd") & "'" _
                                & ",contact_no = '" & txt_contact.Text _
                                & "',address = '" & txt_address.Text & "',course ='" & txt_course.Text & "', school ='" & txt_school.Text & "',a_from ='" & txt_from.Text & "',a_to ='" & txt_to.Text & "',status ='" & cmb_status.Text & "'" _
                            & "WHERE " _
                                & " teacher_id = '" & txt_oldid.Text & "'"
                Call mysql_select(rs_teacher, sql_string)
                
                Call teacherform.Form_Load
                MsgBox "Teacher's information successfully updated."
                 Call set_datagrid(teacherform.dg_teachers, rs_teacher, _
                                        "SELECT " _
                                            & "teacher_id as Teacher_ID, last_name as Last_Name, first_name as First_Name, middle_name as Middle_Name, Gender as Gender, Bday as Date_Of_Birth, contact_no as Contact_Number,address as Address, course as Course, school as School_Attended, a_from as From_Year, a_to as To_Year, status as Status FROM tbl_teacher")
                                        
                    
                   Call teacherform.Form_Load
                Unload Me
                
         
         End If
            End If
            'Call load_teacher
        Else
             If txt_id.Text = "" Or txt_lastname.Text = "" Or txt_firstname.Text = "" Then
                MsgBox "Please input required fields."
            Else
                If is_duplicate("tbl_teacher", "teacher_id", txt_id.Text) Then
                    MsgBox "Teacher ID already exists."
                    Exit Sub
                Else
                ans = MsgBox("Are you sure you want to update teacher's information?", vbYesNo, "Update Teacher")
                    If ans = vbNo Then
                        Exit Sub
                    Else
                sql_string = "UPDATE " _
                                & "tbl_teacher " _
                            & "SET " _
                                & "teacher_id='" & txt_id.Text & "', last_name = '" & txt_lastname.Text & "'," _
                                & "first_name = '" & txt_firstname.Text & "',middle_name = '" _
                                & txt_middlename.Text & "',gender = '" & cmb_gender.Text & "',bday" _
                                & " = '" & Format(dateBday.Value, "yyyy-mm-dd") & "'" _
                                & ",contact_no = '" & txt_contact.Text _
                                & "',address = '" & txt_address.Text & "',course ='" & txt_course.Text & "', school ='" & txt_school.Text & "',a_from ='" & txt_from.Text & "',a_to ='" & txt_to.Text & "',status ='" & cmb_status.Text & "'" _
                            & "WHERE " _
                                & " teacher_id = '" & txt_oldid.Text & "'"
                Call mysql_select(rs_teacher, sql_string)
                sql_string = "UPDATE " _
                                & "tbl_user " _
                            & "SET " _
                                & "teacher_id='" & txt_id.Text & "'" _
                            & "WHERE " _
                                & " teacher_id = '" & txt_oldid.Text & "'"
                Call mysql_select(teacherinformationform.rs_teacher, sql_string)
                MsgBox "Teacher's information successfully updated."
                 Call set_datagrid(teacherform.dg_teachers, rs_teacher, _
                                        "SELECT " _
                                            & "teacher_id as Teacher_ID, last_name as Last_Name, first_name as First_Name, middle_name as Middle_Name, Gender as Gender, Bday as Date_Of_Birth, contact_no as Contact_Number,address as Address, course as Course, school as School_Attended, a_from as From_Year, a_to as To_Year, status as Status FROM tbl_teacher")
                                        
                    
                   
                 Call teacherform.Form_Load
                Unload Me
            End If
            End If
            'Call load_teacher
        End If
        End If
    End If
    Unload Me
End Sub

Private Sub load_teacher()
    Call set_datagrid(teacherform.dg_teachers, rs_teacher, _
                                        "SELECT " _
                                            & "teacher_id as Teacher_ID, last_name as Last_Name, first_name as First_Name, middle_name as Middle_Name, Gender as Gender, Bday as Date_Of_Birth, contact_no as Contact_Number,address as Address, course as Course, school as School_Attended, a_from as From_Year, a_to as To_Year, status as Status FROM tbl_teacher")
                                        
       Unload Me
                                       
End Sub

Private Sub txt_contact_KeyUp(KeyCode As Integer, Shift As Integer)
     If Not IsNumeric(txt_contact.Text) = True Then
        txt_contact.Text = ""
        MsgBox "Please input numbers only."
    End If
End Sub

Private Sub txt_firstname_KeyUp(KeyCode As Integer, Shift As Integer)
     If IsNumeric(Right(txt_firstname.Text, 1)) = True Then
        txt_firstname.Text = ""
        MsgBox "Number is not allowed."
    End If
End Sub

Private Sub txt_from_KeyUp(KeyCode As Integer, Shift As Integer)
     If Not IsNumeric(txt_from.Text) = True Then
        txt_from.Text = ""
        MsgBox "Please input numbers only."
    End If
End Sub

Private Sub txt_lastname_KeyUp(KeyCode As Integer, Shift As Integer)
     If IsNumeric(Right(txt_lastname.Text, 1)) = True Then
        txt_lastname.Text = ""
        MsgBox "Number is not allowed."
    End If
End Sub

Private Sub txt_middlename_KeyUp(KeyCode As Integer, Shift As Integer)
     If IsNumeric(Right(txt_middlename.Text, 1)) = True Then
        txt_middlename.Text = ""
        MsgBox "Number is not allowed."
    End If
End Sub

Private Sub txt_to_KeyUp(KeyCode As Integer, Shift As Integer)
     If Not IsNumeric(txt_to.Text) = True Then
        txt_to.Text = ""
        MsgBox "Please input numbers only."
    End If
End Sub
