VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form teacherinformationform 
   BackColor       =   &H8000000E&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Teacher Information"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10020
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "teacherinformationform.frx":0000
   ScaleHeight     =   5685
   ScaleWidth      =   10020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Height          =   615
      Left            =   7320
      Picture         =   "teacherinformationform.frx":1B3CE
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Height          =   615
      Left            =   7200
      Picture         =   "teacherinformationform.frx":1C149
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   1680
      Width           =   1095
   End
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
      ItemData        =   "teacherinformationform.frx":1D0EC
      Left            =   6600
      List            =   "teacherinformationform.frx":1D0EE
      TabIndex        =   22
      Top             =   1200
      Width           =   2415
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
      ItemData        =   "teacherinformationform.frx":1D0F0
      Left            =   6600
      List            =   "teacherinformationform.frx":1D0FA
      TabIndex        =   21
      Top             =   600
      Width           =   2415
   End
   Begin VB.TextBox txt_oldid 
      Height          =   375
      Left            =   1440
      TabIndex        =   17
      Top             =   4920
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txt_op 
      Height          =   375
      Left            =   3840
      TabIndex        =   16
      Top             =   4920
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
      ItemData        =   "teacherinformationform.frx":1D10C
      Left            =   1920
      List            =   "teacherinformationform.frx":1D116
      TabIndex        =   6
      Top             =   3360
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
      Format          =   52559873
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
      ItemData        =   "teacherinformationform.frx":1D12D
      Left            =   1920
      List            =   "teacherinformationform.frx":1D137
      TabIndex        =   4
      Top             =   2280
      Width           =   3495
   End
   Begin VB.CommandButton cmd_save 
      Height          =   615
      Left            =   1920
      Picture         =   "teacherinformationform.frx":1D149
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmd_cancel 
      Height          =   615
      Left            =   3840
      Picture         =   "teacherinformationform.frx":1E0EC
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4080
      Width           =   1095
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
   Begin MSDataGridLib.DataGrid dg_sections 
      Height          =   2055
      Left            =   5640
      TabIndex        =   23
      Top             =   2400
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   3625
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
   Begin VB.Label Label8 
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   7320
      TabIndex        =   20
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Level"
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
      Left            =   7440
      TabIndex        =   19
      Top             =   360
      Width           =   1455
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
      TabIndex        =   18
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
      TabIndex        =   15
      Top             =   3480
      Width           =   1455
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
      TabIndex        =   14
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
      TabIndex        =   13
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
      TabIndex        =   12
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
      TabIndex        =   11
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
      TabIndex        =   10
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
      TabIndex        =   9
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
Public rs_teacher_sections As New ADODB.Recordset
Dim sql_string As String
Dim section_list(0 To 100) As Long
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

Private Sub cmb_level_Click()
  Call mysql_select(public_rs, "SELECT section_id, section_name FROM db_form137.tbl_section where lvl_name = '" & cmb_level.Text & "'")
  cmb_section.Clear
  
  Dim index As Integer
  
  index = 0
  
  While Not public_rs.EOF
  
    cmb_section.AddItem (public_rs!section_name)
    section_list(index) = public_rs!SECTION_ID
    index = index + 1
  
    public_rs.MoveNext
  Wend
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
    'txt_contact.Text = ""
    'txt_address.Text = ""
    'txt_course.Text = ""
    'txt_school.Text = ""
    'txt_from.Text = ""
    'txt_to.Text = ""
   cmb_status.Text = "On-Duty"
    txt_op.Text = "add"
    'photo.Picture = LoadPicture(App.Path & "\images\photo_teacher\noimage.jpg")
    Unload Me
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
                        & "gender,bday," _
                        & " status)" _
                    & " VALUES (" _
                        & "'" & txt_id.Text & "','" & txt_lastname.Text & "','" _
                        & txt_firstname.Text & "','" & txt_middlename.Text & "','" _
                        & cmb_gender.Text & "','" & Format(dateBday.Value, "yyyy-mm-dd") & "','" _
                        & "On-Duty') "
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
        
        MsgBox "Teacher's information successfully added."
         Call set_datagrid(teacherform.dg_teachers, rs_teacher, _
                                        "SELECT " _
                                            & "teacher_id as Teacher_ID, last_name as Last_Name, first_name as First_Name, middle_name as Middle_Name, Gender as Gender, Bday as Date_Of_Birth, contact_no as Contact_Number,address as Address, course as Course, school as School_Attended, a_from as From_Year, a_to as To_Year, status as Status FROM tbl_teacher")
                                        
        Call teacherform.Form_Load
        Unload Me
        
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
                                & ",status ='" & cmb_status.Text & "'" _
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

Private Sub Command1_Click()
   If (cmb_level.Text = "" Or cmb_section = "") Then
     MsgBox "Please select level and section", vbCritical
     Exit Sub
   End If
   
   Dim sectionID As Long
   
   sectionID = section_list(cmb_section.ListIndex)
   
   Call mysql_select(public_rs, "SELECT * FROM db_form137.tbl_teacher_sections where section_id = " & sectionID)
   
   If (public_rs.RecordCount > 0) Then
     MsgBox "Section is already assigned to a teacher", vbInformation
     Exit Sub
   End If
   

   Call mysql_select(public_rs, "SELECT * FROM db_form137.tbl_teacher_sections where 1 = 2")
   public_rs.AddNew
   public_rs!TEACHER_ID = txt_id
   public_rs!SECTION_ID = sectionID
   public_rs.Update
   MsgBox "Section added to teacher "
  
   Call reloadSectionDataGrid
  
End Sub
Public Sub reloadSectionDataGrid()
   Call mysql_select(rs_teacher_sections, "SELECT lvl_name as Level, section_name as Section, a.id FROM tbl_teacher_sections a, tbl_section b where  a.section_id = b.section_id and a.teacher_id = '" & txt_id & "'")
   Set dg_sections.DataSource = teacherinformationform.rs_teacher_sections
  
   Call formatSectionsDataGrid
End Sub
Public Sub formatSectionsDataGrid()
  dg_sections.Columns(0).Width = 2000
  dg_sections.Columns(1).Width = 2000
   dg_sections.Columns(2).Visible = False
End Sub

Private Sub Command2_Click()
  If (rs_teacher_sections.RecordCount > 0) Then
    Dim response As String
    response = MsgBox("Are you sure you want to delete the record?", vbYesNo, "Question")
    If (response = vbYes) Then
       Call mysql_select(public_rs, "delete from tbl_teacher_sections where id = " & rs_teacher_sections!id)
      MsgBox "Record Deleted", vbInformation
      Call reloadSectionDataGrid
    End If
  End If
End Sub

Private Sub Form_Load()
  Call populateLOV
End Sub

Private Sub populateLOV()

  Call mysql_select(public_rs, "SELECT lvl_name FROM db_form137.tbl_level")
  
  cmb_level.Clear
  
  While (Not public_rs.EOF)
  
    cmb_level.AddItem (public_rs!lvl_name)
    public_rs.MoveNext
  
  Wend
  
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
