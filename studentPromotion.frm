VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form studentPromotion 
   BackColor       =   &H8000000E&
   ClientHeight    =   8340
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13050
   LinkTopic       =   "Form1"
   ScaleHeight     =   8340
   ScaleWidth      =   13050
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmb_adviser 
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
      ItemData        =   "studentPromotion.frx":0000
      Left            =   7440
      List            =   "studentPromotion.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   120
      Width           =   2895
   End
   Begin VB.ComboBox cmb_sy 
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
      ItemData        =   "studentPromotion.frx":0004
      Left            =   2520
      List            =   "studentPromotion.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton cmb_export 
      Caption         =   "Export"
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
      Left            =   10680
      TabIndex        =   2
      Top             =   600
      Width           =   1215
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
      ItemData        =   "studentPromotion.frx":0008
      Left            =   2520
      List            =   "studentPromotion.frx":000A
      TabIndex        =   0
      Top             =   600
      Width           =   2415
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
      ItemData        =   "studentPromotion.frx":000C
      Left            =   7440
      List            =   "studentPromotion.frx":000E
      TabIndex        =   1
      Top             =   600
      Width           =   2895
   End
   Begin MSFlexGridLib.MSFlexGrid flexPromotionBoys 
      Height          =   2895
      Left            =   0
      TabIndex        =   5
      Top             =   1560
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   5106
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid flexPromotionGirls 
      Height          =   3015
      Left            =   0
      TabIndex        =   7
      Top             =   5160
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   5318
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Adviser"
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
      Left            =   6000
      TabIndex        =   12
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "School Year"
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
      Left            =   1080
      TabIndex        =   10
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "Girls"
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
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   4800
      Width           =   13095
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "Boys"
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
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   1200
      Width           =   13095
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Level"
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
      Left            =   1080
      TabIndex        =   4
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Section"
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
      Left            =   5880
      TabIndex        =   3
      Top             =   720
      Width           =   1815
   End
End
Attribute VB_Name = "studentPromotion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public isAdminMode As Boolean

Private teacher_id As String
Private prom_girls_list() As Variant
Private prom_boys_list() As Variant
Private temp_rs As New ADODB.Recordset
Private teacher_name
Private school_year As String
Private teach_list() As Variant

Private Const name_index = 0
Private Const addess_index = 5
Private Const year_in_school_index = 8
Private Const age_index = 9
Private Const days_in_school_index = 10
Private Const grade_index = 11
Private Const action_index = 12

Private Sub cmb_adviser_Click()
  If (isAdminMode And cmb_adviser.ListIndex <> -1) Then
    teacher_id = teach_list(cmb_adviser.ListIndex)
    teacher_name = cmb_adviser.Text
     Dim sqlQuery As String
        sqlQuery = "SELECT distinct b.lvl_name  " & _
                   "FROM tbl_section b, tbl_teacher_sections a " & _
                   "WHERE a.teacher_id='" & teacher_id & "'" & _
                   "      and a.section_id = b.section_id " & _
                   "ORDER BY b.lvl_name ASC "
    Call mysql_select(public_rs, sqlQuery)
    cmb_level.Clear
    cmb_section.Clear
    While Not public_rs.EOF
      cmb_level.AddItem public_rs!lvl_name
      public_rs.MoveNext
    Wend
  End If
End Sub

Private Sub cmb_export_Click()

  Dim excelApp As New Excel.Application
  Dim oBook As New Excel.Workbook
  Dim oSheet As New Excel.Worksheet
  
  Set excelApp = CreateObject("Excel.Application")
  Set oBook = excelApp.Workbooks.Open(CommonHelper.getTemplatesPath & "\Student_Promotion.xlsx")
  Set oSheet = excelApp.Worksheets(1)
  
  excelApp.DisplayAlerts = False
  oBook.SaveAs CommonHelper.getTempPath & "\tmp.xlsx"
  Dim total_age As Double
  Dim average_age As Double
  Dim index As Integer
  
  Dim age_total As Double
  Dim ave_totel As Double
  
  age_total = 0
  ave_totel = 0
    
  On Error GoTo openExcel
  
 
  
  oSheet.Range("C7").value = "SCHOOL YEAR " & school_year
  oSheet.Range("M9").value = Now
  oSheet.Range("M10").value = teacher_name
  oSheet.Range("J10").value = cmb_level.Text
  
  Dim currIndex As Integer
  currIndex = 16
  
  oSheet.Range("A" & currIndex & ":N" & (currIndex + (UBound(prom_boys_list) - 1))).value = prom_boys_list
  
  currIndex = currIndex + UBound(prom_boys_list) + 1
  
  oSheet.Range("A" & currIndex).value = "Total Age"
  oSheet.Range("A" & currIndex).Font.Bold = True
  oSheet.Range("A" & currIndex).HorizontalAlignment = xlCenter
  
  total_age = 0
  
  For index = 1 To UBound(prom_boys_list)
    total_age = total_age + prom_boys_list(index, age_index)
  Next index
  oSheet.Range("J" & currIndex).value = total_age
  oSheet.Range("J" & currIndex).Font.Bold = True
  
  age_total = age_total + total_age
  
  currIndex = currIndex + 1
  
  oSheet.Range("A" & currIndex).Font.Bold = True
  oSheet.Range("A" & currIndex).value = "Average Age"
  oSheet.Range("A" & currIndex).HorizontalAlignment = xlCenter
  
  average_age = Round(total_age / UBound(prom_boys_list), 2)
  ave_totel = ave_totel + average_age
  oSheet.Range("J" & currIndex).value = average_age
  oSheet.Range("J" & currIndex).Font.Bold = True
   
  currIndex = currIndex + 2
  
  oSheet.Range("A" & currIndex).value = "Girls"
  oSheet.Range("A" & currIndex).HorizontalAlignment = xlCenter
  
  currIndex = currIndex + 1
  oSheet.Range("A" & currIndex & ":N" & (currIndex + (UBound(prom_girls_list) - 1))).value = prom_girls_list
  
  currIndex = currIndex + UBound(prom_girls_list) + 1
  
  oSheet.Range("A" & currIndex).value = "Total Age"
  oSheet.Range("A" & currIndex).Font.Bold = True
  oSheet.Range("A" & currIndex).HorizontalAlignment = xlCenter
  
  total_age = 0
  
  For index = 1 To UBound(prom_girls_list)
    total_age = total_age + prom_girls_list(index, age_index)
  Next index
  
  age_total = age_total + total_age
  
  oSheet.Range("J" & currIndex).value = total_age
  oSheet.Range("J" & currIndex).Font.Bold = True
  
  currIndex = currIndex + 1
  
  oSheet.Range("A" & currIndex).Font.Bold = True
  oSheet.Range("A" & currIndex).value = "Average Age"
  oSheet.Range("A" & currIndex).HorizontalAlignment = xlCenter
  
  average_age = Round(total_age / UBound(prom_girls_list), 2)
  ave_totel = ave_totel + average_age
  
  oSheet.Range("J" & currIndex).value = average_age
  oSheet.Range("J" & currIndex).Font.Bold = True
  
  Const max_row_index As Integer = 135
  currIndex = currIndex + 4
  
  oSheet.Range("A" & currIndex & ":A" & max_row_index).EntireRow.Hidden = True
  
  oSheet.Range("A137").value = "Total Age"
  oSheet.Range("A137").Font.Bold = True
  oSheet.Range("A137").HorizontalAlignment = xlCenter
  oSheet.Range("J137").value = age_total
  oSheet.Range("J137").Font.Bold = True
  
  oSheet.Range("A138").Font.Bold = True
  oSheet.Range("A138").value = "Average Age"
  oSheet.Range("A138").HorizontalAlignment = xlCenter
  oSheet.Range("J138").value = Round((ave_totel / 2), 2)
   oSheet.Range("J138").Font.Bold = True
  
  
  
openExcel:
  oSheet.Protect ("password")
  excelApp.Visible = True
  
End Sub

Private Sub cmb_level_Click()
    Dim sqlQuery As String
        
    sqlQuery = "SELECT  Distinct b.section_name " & _
               "FROM tbl_section b, tbl_teacher_sections a " & _
               "WHERE a.teacher_id='" & teacher_id & "'" & _
               "      and a.section_id = b.section_id " & _
               "      and b.lvl_name = '" & cmb_level.Text & "' " & _
               "ORDER BY b.lvl_name ASC "
         
    Call mysql_select(public_rs, sqlQuery)
    
    cmb_section.Clear
    
    While Not public_rs.EOF
      cmb_section.AddItem public_rs!section_name
      public_rs.MoveNext
    Wend
    
End Sub

Private Sub cmb_section_Click()
  If cmb_section.Text <> vbNullString Then
    Dim base_query As String
    
    base_query = "Select stud.student_id, stud.last_name, stud.first_name, stud.middle_name, stud.address " & _
                 "       , ( " & _
                 "            select count(*) " & _
                 "            from tbl_student a, tbl_student_level b " & _
                 "            Where b.id = a.student_id" & _
                 "                  and a.student_id = stud.student_id " & _
                 "            group by a.student_id " & _
                 "          ) as years_in_school " & _
                 "       , TRUNCATE(FLOOR(((12 * (YEAR(NOW())- YEAR(bday))+ (MONTH(NOW())- MONTH(stud.bday))) / 12) * 4) / 4 , 2) AS Age " & _
                 "       , ( " & _
                 "            select a.no_days_present " & _
                 "            from tbl_attendance a " & _
                 "            where a.id = stud.student_id " & _
                 "                  and a.SY = '" & school_year & "' " & _
                 "            limit 1 " & _
                 "          ) as days_in_school " & _
                 "from tbl_student stud, tbl_student_level lvlsec " & _
                 "Where lvlsec.id = stud.student_id " & _
                 "      and lvlsec.lvl_name = '" & cmb_level.Text & "' " & _
                 "      and lvlsec.section_name = '" & cmb_section.Text & "' " & _
                 "      and lvlsec.sy = '" & tbl_student_level & "' "
                 
    'InputBox "", "", base_query
    
    Call mysql_select(public_rs, base_query & " and stud.gender = 'Male' ")
    If (public_rs.RecordCount > 0) Then
      prom_boys_list = populatePromotionFlex(flexPromotionBoys, public_rs)
    End If
    
    Call mysql_select(public_rs, base_query & " and stud.gender = 'Female' ")
    If (public_rs.RecordCount > 0) Then
      prom_girls_list = populatePromotionFlex(flexPromotionGirls, public_rs)
    End If

  End If
End Sub
Public Function populatePromotionFlex(flexGrid As MSFlexGrid, rs As ADODB.Recordset) As Variant()
  Dim index As String
  index = 1
  Dim prom_list() As Variant
  ReDim prom_list(1 To rs.RecordCount, 0 To 13) As Variant
  
  With flexGrid
    .Clear
    
    .Rows = rs.RecordCount + 1
    .Cols = 8
    .WordWrap = True
    
    .RowHeight(0) = 1050
    .ColAlignment(0) = flexAlignLeftCenter
    .ColWidth(0) = 3700
    .ColAlignment(1) = flexAlignCenterCenter
    .ColWidth(1) = 2000
    .ColAlignment(2) = flexAlignCenterCenter
    .ColWidth(2) = 900
    .ColAlignment(3) = flexAlignCenterCenter
    .ColWidth(4) = 900
    .ColAlignment(4) = flexAlignCenterCenter
    .ColWidth(4) = 900
    .ColAlignment(5) = flexAlignCenterCenter
    .ColWidth(5) = 900
    .ColAlignment(6) = flexAlignCenterCenter
    .ColWidth(6) = 900
    .ColAlignment(7) = flexAlignCenterCenter
    .ColWidth(7) = 2250
    
    
    .TextMatrix(0, 0) = "                         NAME"
    .TextMatrix(0, 1) = "HOME ADDRESS"
    .TextMatrix(0, 2) = "YEARS IN SHCOOL"
    .TextMatrix(0, 3) = "AGE"
    .TextMatrix(0, 4) = "TOTAL NUMBER OF DAYS IN GRADE"
    .TextMatrix(0, 5) = "FINAL RATING"
    .TextMatrix(0, 6) = "ACTION TAKEN"
    .TextMatrix(0, 7) = "REMARK"
    
    Dim sql_query As String
    Dim divider As Integer
    Dim total_grade As Integer
    While Not rs.EOF
      
      prom_list(index, name_index) = index & " " & CommonHelper.extractStringValue(rs!LAST_NAME) & ", " & CommonHelper.extractStringValue(rs!FIRST_NAME) & " " & toIntial(rs!middle_name)
      .TextMatrix(index, 0) = prom_list(index, name_index)
      
      prom_list(index, addess_index) = CommonHelper.extractStringValue(rs!address)
      .TextMatrix(index, 1) = prom_list(index, addess_index)
      
      prom_list(index, year_in_school_index) = CommonHelper.extractStringValue(rs!years_in_school)
      .TextMatrix(index, 2) = prom_list(index, year_in_school_index)
      
      prom_list(index, age_index) = CommonHelper.extractStringValue(rs!age)
      .TextMatrix(index, 3) = prom_list(index, age_index)
      
      prom_list(index, days_in_school_index) = CommonHelper.extractStringValue(rs!days_in_school)
      .TextMatrix(index, 4) = prom_list(index, days_in_school_index)
    
      sql_query = "Select GRADE " & _
                  "From tbl_grade " & _
                  "Where SY = '" & school_year & "' " & _
                  "      And ID = '" & rs!student_id & "' "

      divider = 0
      Call mysql_select(temp_rs, sql_query)
      
      total_grade = 0
      
      While Not temp_rs.EOF
        If (val(temp_rs!grade) > 0) Then
          total_grade = total_grade + val(temp_rs!grade)
          divider = divider + 1
        End If
        temp_rs.MoveNext
      Wend
      
      If (total_grade > 0) Then
      
         prom_list(index, grade_index) = Round(val(total_grade / divider), 0)
        .TextMatrix(index, 5) = prom_list(index, grade_index)
        
        If (val(prom_list(index, 11)) >= 75) Then
          prom_list(index, action_index) = "Prom."
        Else
          prom_list(index, action_index) = "Failed"
        End If
        
        .TextMatrix(index, 6) = prom_list(index, action_index)
        prom_list(index, 13) = ""
      
      End If
      
      index = index + 1
      rs.MoveNext
    Wend
        
  End With
  
  populatePromotionFlex = prom_list()
  
End Function
Private Function toIntial(m_name As String)
  If (m_name <> vbNullString) Then
    toIntial = UCase(Right(m_name, 1)) & "."
  Else
    toIntial = ""
  End If
End Function

Private Sub cmb_sy_Click()
  If (isAdminMode) Then
    school_year = cmb_sy.Text
    cmb_adviser.ListIndex = -1
    cmb_level.Clear
    cmb_section.Clear
  End If
End Sub

Private Sub Form_Load()
    If (isAdminMode = False) Then
      Call mysql_select(public_rs, "SELECT * FROM tbl_user WHERE Username = '" & mainform.lbl_username.Caption & "'")
      teacher_id = public_rs.Fields("ID")
      school_year = mainteacherform.cmb_sy.Text
    
      Dim sqlQuery As String
        
      sqlQuery = "SELECT  Distinct b.lvl_name " & _
                 "FROM tbl_section b, tbl_teacher_sections a " & _
                 "WHERE a.teacher_id='" & teacher_id & "'" & _
                 "      and a.section_id = b.section_id " & _
                 "ORDER BY b.lvl_name ASC "
         
      Call mysql_select(public_rs, sqlQuery)
    
      cmb_level.Clear
      While Not public_rs.EOF
        cmb_level.AddItem public_rs!lvl_name
        public_rs.MoveNext
      Wend
      
      cmb_sy.Clear
      cmb_sy.AddItem school_year
      cmb_sy.Text = school_year
      cmb_sy.Enabled = False
      
      Call mysql_select(public_rs, "SELECT * FROM tbl_teacher where teacher_id = '" & teacher_id & "'")
      teacher_name = public_rs!FIRST_NAME & " " & public_rs!middle_name & " " & public_rs!LAST_NAME
      
      cmb_adviser.Clear
      cmb_adviser.Enabled = False
      cmb_adviser.AddItem teacher_name
      cmb_adviser.Text = teacher_name
    Else
      
      cmb_sy.Clear
      cmb_sy.Enabled = True
      Call mysql_select(public_rs, "SELECT distinct sy FROM tbl_student_level")
      
      If (public_rs.RecordCount > 0) Then
        While Not public_rs.EOF
          cmb_sy.AddItem public_rs!sy
          public_rs.MoveNext
        Wend
        cmb_sy.ListIndex = 0
        school_year = cmb_sy.Text
      Else
         school_year = ""
      End If
      
      cmb_adviser.Clear
      cmb_adviser.Enabled = True
      Call mysql_select(public_rs, "SELECT teacher_id, concat(First_Name, ' ', middle_name, ' ', LAST_NAME) as name FROM tbl_teacher")
      
      ReDim teach_list(0 To public_rs.RecordCount) As Variant
      Dim index As Integer
      index = 0
      While Not public_rs.EOF
        teach_list(index) = public_rs!teacher_id
        cmb_adviser.AddItem public_rs!Name
        index = index + 1
        public_rs.MoveNext
      Wend
           
    End If
End Sub

Private Sub MSFlexGrid1_Click()

End Sub

