VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form promotionform 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Student Promotion Status"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "promotionform.frx":0000
   ScaleHeight     =   6000
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Top             =   -120
      Width           =   8775
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
         ItemData        =   "promotionform.frx":1B3CE
         Left            =   6120
         List            =   "promotionform.frx":1B3D8
         TabIndex        =   2
         Top             =   480
         Width           =   2535
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
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   2535
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
         Left            =   2760
         TabIndex        =   1
         Top             =   480
         Width           =   3255
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Choose Gender"
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
         Left            =   6120
         TabIndex        =   8
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Choose Section:"
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
         Left            =   2760
         TabIndex        =   6
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Choose Level:"
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
         TabIndex        =   5
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmd_print 
      Height          =   615
      Left            =   3720
      Picture         =   "promotionform.frx":1B3E9
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5280
      Width           =   1575
   End
   Begin MSDataGridLib.DataGrid dg_promotion 
      Height          =   4215
      Left            =   240
      TabIndex        =   7
      Top             =   960
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   7435
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
Attribute VB_Name = "promotionform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rs_promotion As New ADODB.Recordset
Public rs_years As New ADODB.Recordset
Public rs_age As New ADODB.Recordset
Public rs_days As New ADODB.Recordset
Public rs_rating As New ADODB.Recordset

Dim sql_string As String
Dim gender As String
Private Sub cmb_gender_Click()
    
    If cmb_gender.Text = "Boys" Then
        gender = "Male"
    Else
        gender = "Female"
    End If
    Call set_datagrid(dg_promotion, rs_promotion, _
                                        "SELECT " _
                                            & " Name, ID, Address, Years_in_School, Age, Number_of_Days, Grade_Remark,Final_Rating, Action_Taken, Remark FROM tbl_promotion WHERE Section = '" & cmb_section.Text & "' AND Level='" & cmb_level.Text & "'AND Gender='" & gender & "' ORDER BY Name ASC")
    
    
End Sub

Private Sub cmb_gender_KeyUp(KeyCode As Integer, Shift As Integer)
    MsgBox "Please select gender from the list."
    cmb_gender.Text = ""
End Sub

Private Sub cmb_level_Change()
     Call mysql_select(public_rs, "SELECT * FROM tbl_section WHERE SY='" & mainform.lbl_sy.Caption & "'AND lvl_name='" & cmb_level.Text & "'")
    promotionform.cmb_section.Clear
    While Not public_rs.EOF
        promotionform.cmb_section.AddItem (public_rs.Fields("section_name").Value)
        public_rs.MoveNext
    Wend
End Sub

Private Sub cmb_level_Click()
     Call mysql_select(public_rs, "SELECT * FROM tbl_section WHERE lvl_name='" & cmb_level.Text & "'")
    promotionform.cmb_section.Clear
    While Not public_rs.EOF
        promotionform.cmb_section.AddItem (public_rs.Fields("section_name").Value)
        public_rs.MoveNext
    Wend
End Sub

Private Sub cmb_level_KeyUp(KeyCode As Integer, Shift As Integer)
    MsgBox "Please select level from the list."
    cmb_level.Text = ""
End Sub

Private Sub cmb_section_Click()
    cmb_gender.Text = ""
    Dim stud_name, id, gender, address, years, age, days, rating, remark, remark2 As String
    Dim average As Double
    Dim no_subj As Integer
    Call set_datagrid(dg_promotion, rs_promotion, _
                                        "SELECT " _
                                            & "a.student_id as LRN, a.last_name as Last_Name, a.First_Name,a.Middle_Name FROM tbl_student a LEFT JOIN tbl_student_level b ON a.student_id = b.ID WHERE b.section_name = '" & cmb_section.Text & "' ORDER BY a.Last_Name ASC")
    
    
     Call mysql_select(public_rs, "DELETE  FROM tbl_promotion  WHERE Section = '" & cmb_section.Text & "' AND Level='" & cmb_level.Text & "' ")
    Call mysql_select(public_rs, "SELECT a.* FROM tbl_student a LEFT JOIN tbl_student_level b ON a.student_id = b.ID WHERE b.section_name = '" & cmb_section.Text & "' ORDER BY a.Last_Name ASC")
    If public_rs.RecordCount = 0 Then
        MsgBox "No students enrolled in this section."
        Exit Sub
    End If
    While Not public_rs.EOF
            stud_name = public_rs.Fields("last_name") & ", " & public_rs.Fields("first_name").Value & " " & public_rs.Fields("middle_name").Value
            id = public_rs.Fields("student_id").Value
            gender = public_rs.Fields("Gender").Value
            address = public_rs.Fields("Address").Value
          Call mysql_select(rs_years, "SELECT * FROM tbl_student_level  WHERE ID='" & id & " '")
          If rs_years.RecordCount = 0 Then
            years = "0"
        Else
            years = Str(rs_years.RecordCount)
          End If
          Call mysql_select(rs_age, "SELECT TRUNCATE(FLOOR(((12 * (YEAR(NOW())- YEAR(bday))+ (MONTH(NOW())- MONTH( bday))) / 12) * 4) / 4 , 2) AS Age From tbl_student WHERE student_id ='" & id & "'")
            If rs_age.RecordCount = 0 Then
                age = "0"
            Else
                age = Str(rs_age.Fields("Age").Value)
              End If
         Call mysql_select(rs_days, "SELECT * FROM tbl_attendance  WHERE ID='" & id & "' ")
          If rs_days.RecordCount = 0 Then
            days = "0"
        Else
            days = Str(rs_days.Fields("no_days_present").Value)
          End If
           Call mysql_select(rs_rating, "SELECT subject_code as Subject, period, grade, remark FROM tbl_grade WHERE ID = '" & id & "' AND Period='Final'")
        no_subj = rs_rating.RecordCount
        If no_subj = 0 Then
            rating = "No grade"
        Else
            average = 0
            While Not rs_rating.EOF
                average = val(rs_rating.Fields("grade")) + average
                rs_rating.MoveNext
            Wend
            average = average / no_subj
            average = Round(average, 2)
            rating = Str(average)
            If average >= 90 Then
              remark = "A"
              ElseIf average >= 85 Then
              remark = "P"
              ElseIf average >= 80 Then
                  remark = "AP"
              ElseIf average >= 74 Then
                  remark = "D"
              Else
                  remark = "B"
            End If
        End If
        If years = "0" Or age = "0" Or days = "0" Or rating = "No grade" Or remark = "B" Then
            remark2 = "Incomplete"
        Else
            remark2 = "Promote"
        End If
        
         sql_string = "INSERT INTO " _
                        & "tbl_promotion (Name,ID,Gender," _
                        & "Level,Section,Address,Years_in_School," _
                        & "Age,Number_of_Days,Grade_Remark,Final_Rating,Action_Taken, Remark)" _
                    & " VALUES (" _
                        & "'" & stud_name & "','" & id & "','" _
                        & gender & "', '" _
                        & cmb_level.Text & "','" & cmb_section.Text & "','" _
                        & address & "','" _
                        & years & "', '" & age & "','" & days & "','" & remark & "','" & rating & "','','" & remark & "')"
        Call mysql_select(rs_promotion, sql_string)
        public_rs.MoveNext
    Wend


End Sub

Private Sub cmb_section_KeyUp(KeyCode As Integer, Shift As Integer)
    MsgBox "Please select section from the list."
    cmb_section.Text = ""
End Sub

Private Sub cmd_print_Click()
Dim id, teacher As String
    If cmb_gender.Text = "" Then
        MsgBox "Please select gender first."
        Exit Sub
    End If

     
    If dg_promotion.DataSource Is Nothing Then
        MsgBox "No record."
        cmb_gender.Text = ""
        Exit Sub
    End If
  
     If rs_promotion.RecordCount = 0 Then
        MsgBox "No record."
       Exit Sub
    End If
    
    'dr_promotion.Sections(2).Controls("lbl_sy").Caption = mainform.lbl_sy.Caption
      dr_promotion.Sections(2).Controls("lbl_date").Caption = Format(Now, "mmmm dd, yyyy") & "/ " & Time
        dr_promotion.Sections(2).Controls("lbl_level").Caption = cmb_level.Text
        dr_promotion.Sections(2).Controls("lbl_section").Caption = cmb_section.Text
        Call mysql_select(public_rs, "SELECT * FROM tbl_section")
        If public_rs.RecordCount = 0 Then
            teacher = ""
        Else
            id = public_rs.Fields("teacher_id").Value
        End If
          Call mysql_select(public_rs, "SELECT CONCAT(CONCAT(first_name,' '),last_name) as Name FROM tbl_teacher WHERE teacher_id='" & id & "'")
            If public_rs.RecordCount = 0 Then
                teacher = ""
                Else
                   teacher = public_rs.Fields("Name").Value
            End If
            Dim total1, total2, ave1, ave2 As Double
            Dim no As Integer
            Call mysql_select(public_rs, "SELECT Age FROM tbl_promotion WHERE  Section='" & cmb_section.Text & "' AND Level='" & cmb_level.Text & "'AND Gender='" & gender & "'")
        no = public_rs.RecordCount
        If no = 0 Then
            total1 = 0
            ave1 = 0
        Else
            total1 = 0
            ave1 = 0
            While Not public_rs.EOF
                total1 = val(public_rs.Fields("Age")) + total1
                public_rs.MoveNext
            Wend
            ave1 = total1 / no
            ave1 = Round(ave1, 2)
        End If
        Call mysql_select(public_rs, "SELECT Age FROM tbl_promotion WHERE  Section='" & cmb_section.Text & "' AND Level='" & cmb_level.Text & "'")
        no = public_rs.RecordCount
        If no = 0 Then
            total2 = 0
            ave2 = 0
        Else
            total2 = 0
            ave2 = 0
            While Not public_rs.EOF
                total2 = val(public_rs.Fields("Age")) + total2
                public_rs.MoveNext
            Wend
            ave2 = total2 / no
            ave2 = Round(ave2, 2)
        End If
        dr_promotion.Sections(2).Controls("lbl_teacher").Caption = teacher
        dr_promotion.Sections(5).Controls("lbl_total").Caption = total1
        dr_promotion.Sections(5).Controls("lbl_average").Caption = ave1
        dr_promotion.Sections(5).Controls("lbl_total2").Caption = total2
        dr_promotion.Sections(5).Controls("lbl_average2").Caption = ave2
         Set dr_promotion.DataSource = rs_promotion
        dr_promotion.Show vbModal, Me
End Sub

Private Sub Form_Load()
    Call mysql_select(public_rs, "SELECT * FROM tbl_level ")
    promotionform.cmb_level.Clear
    While Not public_rs.EOF
        promotionform.cmb_level.AddItem (public_rs.Fields("lvl_name").Value)
        public_rs.MoveNext
    Wend
     Call load_form(promotionform, True)
End Sub

